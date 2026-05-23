# 💰 Finanz-Tracker

Persönliche Finanz-App · PWA (installierbar auf iPhone, Android & Desktop) · Darkmode · KI-Kategorisierung · Bankanbindung

---

## 🚀 Schritt 1 — Lokal starten (5 Minuten)

**Voraussetzung:** [Node.js](https://nodejs.org) installieren (LTS-Version)

```bash
# 1. In den Projektordner wechseln
cd finance-tracker

# 2. Abhängigkeiten installieren
npm install

# 3. App starten
npm run dev
```

Öffne dann http://localhost:5173 im Browser.

---

## 🌐 Schritt 2 — Online deployen mit Vercel (kostenlos, 10 Minuten)

### 2a. GitHub-Repository erstellen
1. Gehe zu [github.com](https://github.com) → „New repository"
2. Name: `finance-tracker` → „Create repository"
3. Führe die angezeigten Befehle in deinem Projektordner aus:

```bash
git init
git add .
git commit -m "Finanz-Tracker initial"
git branch -M main
git remote add origin https://github.com/DEIN-USERNAME/finance-tracker.git
git push -u origin main
```

### 2b. Vercel verbinden
1. Gehe zu [vercel.com](https://vercel.com) → „Sign up with GitHub"
2. „New Project" → dein `finance-tracker` Repository auswählen
3. Alles so lassen (Vite wird automatisch erkannt) → „Deploy"
4. ✅ Nach 60 Sekunden hast du eine URL wie `finance-tracker-xyz.vercel.app`

**Diese URL funktioniert auf jedem Gerät — Desktop, iPhone, Android.**

---

## 📱 Schritt 3 — Als App installieren (PWA)

### Auf iPhone (Safari)
1. finance-tracker-xyz.vercel.app in Safari öffnen
2. Teilen-Symbol (□↑) antippen → „Zum Home-Bildschirm"
3. „Hinzufügen" → App erscheint auf dem Home-Bildschirm

### Auf Android (Chrome)
1. URL in Chrome öffnen
2. Banner erscheint automatisch: „App installieren"
3. Alternativ: Menü (⋮) → „App installieren"

### Auf Desktop (Chrome/Edge)
1. URL öffnen
2. Adressleiste: Symbol „App installieren" (⊕) klicken

---

## 🏦 Schritt 4 — Echtzeit-Bankanbindung (GoCardless)

GoCardless Open Banking (früher Nordigen) ist für Privatpersonen **kostenlos** und unterstützt alle deutschen Banken via PSD2.

**Unterstützte Banken:** DKB, ING, Sparkasse, Volksbank, Deutsche Bank, Commerzbank, Comdirect, N26, Revolut, und 2.000+ weitere.

### 4a. GoCardless-Account erstellen
1. Gehe zu [bankaccountdata.gocardless.com](https://bankaccountdata.gocardless.com)
2. „Sign up free" → Konto erstellen
3. Im Dashboard: API → „Create new API key"
4. Notiere `SECRET_ID` und `SECRET_KEY`

### 4b. Backend mit Supabase aufsetzen (kostenlos)

**Warum Backend?** Die API-Keys dürfen nicht im Browser stehen (Sicherheit).

1. Gehe zu [supabase.com](https://supabase.com) → „New Project"
2. Im Supabase-Dashboard → „Edge Functions" → „New Function" → Name: `bank-sync`
3. Füge folgenden Code ein:

```typescript
// supabase/functions/bank-sync/index.ts
import { serve } from "https://deno.land/std@0.168.0/http/server.ts"

const GC_BASE = "https://bankaccountdata.gocardless.com/api/v2"
const SECRET_ID  = Deno.env.get("GC_SECRET_ID")!
const SECRET_KEY = Deno.env.get("GC_SECRET_KEY")!

serve(async (req) => {
  const { action, ...params } = await req.json()
  const headers = { "Content-Type": "application/json" }

  // Token holen
  const tokenRes = await fetch(`${GC_BASE}/token/new/`, {
    method: "POST", headers,
    body: JSON.stringify({ secret_id: SECRET_ID, secret_key: SECRET_KEY })
  })
  const { access } = await tokenRes.json()
  const authH = { ...headers, "Authorization": `Bearer ${access}` }

  if (action === "list_banks") {
    // Verfügbare Banken für Deutschland laden
    const r = await fetch(`${GC_BASE}/institutions/?country=de`, { headers: authH })
    return new Response(await r.text(), { headers })
  }

  if (action === "connect_bank") {
    // Bank-Verbindung initiieren (öffnet Bank-Login)
    const r = await fetch(`${GC_BASE}/requisitions/`, {
      method: "POST", headers: authH,
      body: JSON.stringify({
        redirect: params.redirect_url,
        institution_id: params.institution_id,
        reference: "finanztracker-" + Date.now()
      })
    })
    return new Response(await r.text(), { headers })
  }

  if (action === "get_transactions") {
    // Transaktionen der letzten 90 Tage abrufen
    const r = await fetch(
      `${GC_BASE}/accounts/${params.account_id}/transactions/?date_from=${params.from}`,
      { headers: authH }
    )
    return new Response(await r.text(), { headers })
  }

  return new Response(JSON.stringify({ error: "Unknown action" }), { status: 400, headers })
})
```

4. Umgebungsvariablen setzen: Supabase → Settings → Edge Functions → Secrets:
   - `GC_SECRET_ID` = dein GoCardless Secret ID
   - `GC_SECRET_KEY` = dein GoCardless Secret Key

### 4c. App mit Bank verbinden

Füge in `src/App.jsx` folgende Funktion hinzu:

```javascript
const SUPABASE_URL = "https://DEIN-PROJEKT.supabase.co/functions/v1/bank-sync"

async function syncBank(accountId) {
  const from = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]
  const res = await fetch(SUPABASE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ action: 'get_transactions', account_id: accountId, from })
  })
  const data = await res.json()
  const txs = data.transactions?.booked || []
  return txs.map(t => ({
    id: Date.now() + Math.random(),
    desc: t.remittanceInformationUnstructured || t.creditorName || 'Bank',
    amount: Math.abs(parseFloat(t.transactionAmount.amount)),
    type: parseFloat(t.transactionAmount.amount) >= 0 ? 'income' : 'expense',
    cat: 'Sonstiges',  // KI kategorisiert automatisch
    date: t.valueDate || t.bookingDate,
    source: 'Bankkonto'
  }))
}
```

---

## 📊 Funktionen

| Feature | Status |
|---|---|
| Dashboard mit Charts | ✅ |
| Buchungen manuell | ✅ |
| KI-Kategorisierung | ✅ |
| Budget-Limits & Warnungen | ✅ |
| Wiederkehrende Buchungen | ✅ |
| CSV-Import | ✅ |
| Excel / CSV / PDF Export | ✅ |
| Darkmode | ✅ |
| Offline-Nutzung (PWA) | ✅ |
| Echtzeit-Bankanbindung | 🔧 Setup erforderlich |
| Multi-Device Sync | 🔧 Mit Supabase möglich |

---

## 🔒 Datenschutz

- Alle Daten liegen **lokal im Browser** (localStorage)
- Kein Tracking, keine Werbung
- Bei Bankanbindung: Daten fließen nur über dein eigenes Supabase-Backend
- KI-Kategorisierung nutzt die Anthropic API (nur Buchungsbeschreibung, anonym)

---

## 🛠 Technologie

- **React 18** + **Vite**
- **Recharts** (Charts)
- **SheetJS** (Excel-Export)
- **vite-plugin-pwa** (Installierbar auf allen Geräten)
- **GoCardless** (Bank-API, PSD2-konform)
- **Supabase** (Backend + Datenbank)
- **Anthropic Claude API** (KI-Kategorisierung)
