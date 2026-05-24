// SparkasseImport.jsx — Automatischer Sparkasse CSV Importer
// Füge diese Komponente in App.jsx ein (vor der letzten Zeile)

// ── Sparkasse Import ───────────────────────────────────────────────────────

export function SparkasseImport({ setTxs, nid, setNid, toast, onBack, cats, rules, existingTxs }) {
  const [step, setStep]         = useState('upload')   // upload | preview | done
  const [rows, setRows]         = useState([])
  const [preview, setPreview]   = useState([])
  const [stats, setStats]       = useState(null)
  const [importing, setImporting] = useState(false)

  // Sparkasse CSV parsen
  const parseSparkasseCSV = (text) => {
    const lines = text.split('\n').filter(l => l.trim())
    const results = []
    for (let i = 1; i < lines.length; i++) { // Zeile 0 = Header
      const cols = parseCSVLine(lines[i])
      if (!cols || cols.length < 16) continue

      const info      = (cols[16] || '').trim()
      const betragRaw = (cols[14] || '').replace(/"/g,'').trim()
      const betrag    = parseFloat(betragRaw.replace(',','.'))

      // Vorgemerkte und Null-Buchungen überspringen
      if (info === 'Umsatz vorgemerkt') continue
      if (isNaN(betrag) || betrag === 0) continue

      const buchungstext   = (cols[3]  || '').replace(/"/g,'').trim()
      if (buchungstext === 'ABSCHLUSS') continue

      const buchungstag    = (cols[1]  || '').replace(/"/g,'').trim()
      const verwendungszweck = (cols[4] || '').replace(/"/g,'').trim()
      const empfaenger     = (cols[11] || '').replace(/"/g,'').trim()
      const kundenreferenz = (cols[7]  || '').replace(/"/g,'').trim()

      // Datum konvertieren DD.MM.YY → YYYY-MM-DD
      const date = parseDatum(buchungstag)
      if (!date) continue

      // Beschreibung: Empfänger bevorzugen, sonst Verwendungszweck kürzen
      const desc = empfaenger
        ? cleanEmpfaenger(empfaenger)
        : kuerzeVerwendung(verwendungszweck)

      // Typ ermitteln
      const type = betrag >= 0 ? 'income' : 'expense'
      const amount = Math.abs(betrag)

      // Duplikat-Key: Datum + Betrag + Empfänger
      const dupKey = `${date}|${amount}|${desc}`

      results.push({ date, desc, amount, type, source: 'Sparkasse', dupKey, verwendungszweck, empfaenger })
    }
    return results
  }

  const parseCSVLine = (line) => {
    const result = []
    let current = ''
    let inQuotes = false
    for (let i = 0; i < line.length; i++) {
      const ch = line[i]
      if (ch === '"') {
        inQuotes = !inQuotes
      } else if (ch === ';' && !inQuotes) {
        result.push(current)
        current = ''
      } else {
        current += ch
      }
    }
    result.push(current)
    return result
  }

  const parseDatum = (d) => {
    // DD.MM.YY oder DD.MM.YYYY
    const parts = d.split('.')
    if (parts.length !== 3) return null
    const day   = parts[0].padStart(2,'0')
    const month = parts[1].padStart(2,'0')
    let year    = parts[2]
    if (year.length === 2) year = '20' + year
    return `${year}-${month}-${day}`
  }

  const cleanEmpfaenger = (e) => {
    // Lange Adressen kürzen, nur den Namen nehmen
    return e.split(/\s{2,}/)[0].trim().slice(0, 50)
  }

  const kuerzeVerwendung = (v) => {
    // Ersten sinnvollen Teil nehmen
    return v.slice(0, 50).trim()
  }

  const onFile = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return

    // Encoding erkennen (Sparkasse nutzt oft Windows-1252)
    const buffer = await file.arrayBuffer()
    let text
    try {
      text = new TextDecoder('windows-1252').decode(buffer)
    } catch {
      text = new TextDecoder('utf-8').decode(buffer)
    }

    // Sparkasse-Format prüfen
    if (!text.includes('Buchungstag') && !text.includes('Auftragskonto')) {
      toast('Kein Sparkasse-Format erkannt', 'err')
      return
    }

    const parsed = parseSparkasseCSV(text)
    if (parsed.length === 0) {
      toast('Keine Buchungen gefunden', 'err')
      return
    }

    // Duplikate prüfen
    const existingKeys = new Set(
      existingTxs.map(t => `${t.date}|${t.amount}|${t.desc}`)
    )
    const neu      = parsed.filter(r => !existingKeys.has(r.dupKey))
    const duplikate = parsed.filter(r =>  existingKeys.has(r.dupKey))

    setRows(neu)
    setPreview(neu.slice(0, 10))
    setStats({ total: parsed.length, neu: neu.length, duplikate: duplikate.length })
    setStep('preview')
  }

  const doImport = () => {
    if (rows.length === 0) { toast('Keine neuen Buchungen', 'err'); return }
    setImporting(true)
    let id = nid
    const newTxs = rows.map(r => {
      // KI-Kategorisierung: Empfänger + Verwendungszweck kombinieren
      const searchText = `${r.empfaenger} ${r.verwendungszweck}`.toLowerCase()
      const cat = aiCatSparkasse(searchText, rules)
      return { id: id++, desc: r.desc, amount: r.amount, type: r.type, cat, date: r.date, source: 'Sparkasse' }
    })
    setTxs(p => [...newTxs, ...p])
    setNid(id)
    setStats(s => ({ ...s, imported: newTxs.length }))
    setStep('done')
    setImporting(false)
    toast(`${newTxs.length} Buchungen importiert ✓`)
  }

  // Sparkasse-spezifische KI (nutzt kombinierten Suchtext)
  const aiCatSparkasse = (searchText, rules) => {
    for (const [cat, keywords] of Object.entries(rules || {})) {
      const kws = typeof keywords === 'string'
        ? keywords.split(',').map(k => k.trim().toLowerCase())
        : (keywords || [])
      if (kws.some(k => k && searchText.includes(k))) return cat
    }
    return 'Sonstiges'
  }

  return (
    <div className="fade-up">
      <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:'1.25rem' }}>
        <button onClick={onBack} style={{ color:'var(--t2)', fontSize:18 }}>
          <i className="ti ti-arrow-left" />
        </button>
        <span style={{ fontSize:18, fontWeight:600 }}>Sparkasse Import</span>
      </div>

      {step === 'upload' && (
        <div>
          <Card style={{ background:'var(--card2)', marginBottom:'1rem' }}>
            <div style={{ fontSize:13, color:'var(--t2)', lineHeight:1.6 }}>
              <strong style={{ color:'var(--t1)' }}>So lädst du deine CSV herunter:</strong><br />
              1. Sparkasse Online-Banking öffnen<br />
              2. Konto auswählen → Umsätze<br />
              3. Exportieren → CSV-CAMT Format oder CSV<br />
              4. Datei hier hochladen
            </div>
          </Card>

          <Card>
            <div style={{ textAlign:'center', padding:'1.5rem 0' }}>
              <i className="ti ti-file-upload" style={{ fontSize:40, color:'var(--sky)', display:'block', marginBottom:12 }} />
              <div style={{ fontSize:14, color:'var(--t2)', marginBottom:16 }}>
                Sparkasse CSV-Datei hochladen
              </div>
              <label style={{
                display:'inline-flex', alignItems:'center', gap:8,
                padding:'10px 20px', borderRadius:'var(--r-md)',
                background:'var(--t1)', color:'var(--page)', fontSize:14, fontWeight:500, cursor:'pointer'
              }}>
                <i className="ti ti-upload" /> Datei wählen
                <input type="file" accept=".csv,.CSV" onChange={onFile} style={{ display:'none' }} />
              </label>
            </div>
          </Card>
        </div>
      )}

      {step === 'preview' && stats && (
        <div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8, marginBottom:'1rem' }}>
            {[
              { label:'Gesamt',     val:stats.total,     c:'var(--t1)',    bg:'var(--card2)'    },
              { label:'Neu',        val:stats.neu,       c:'var(--mint)',  bg:'var(--mint-bg)'  },
              { label:'Duplikate',  val:stats.duplikate, c:'var(--t3)',    bg:'var(--card3)'    },
            ].map(m => (
              <div key={m.label} style={{ background:m.bg, borderRadius:'var(--r-lg)', padding:'0.8rem' }}>
                <div style={{ fontSize:11, color:m.c, opacity:0.8, marginBottom:4 }}>{m.label}</div>
                <div style={{ fontSize:20, fontWeight:600, color:m.c }}>{m.val}</div>
              </div>
            ))}
          </div>

          <Card>
            <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:10 }}>
              Vorschau (erste {preview.length} Buchungen)
            </div>
            {preview.map((r, i) => {
              const searchText = `${r.empfaenger} ${r.verwendungszweck}`.toLowerCase()
              const cat = aiCatSparkasse(searchText, rules)
              const info = getCatInfo(cat, cats)
              return (
                <div key={i} style={{ display:'flex', alignItems:'center', gap:10, padding:'8px 0', borderTop: i>0?'1px solid var(--brd)':'none' }}>
                  <div style={{ width:32, height:32, borderRadius:'50%', background:info.bg, color:info.c, display:'flex', alignItems:'center', justifyContent:'center', fontSize:14, flexShrink:0 }}>
                    <i className={`ti ${info.icon}`} />
                  </div>
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:13, fontWeight:500, color:'var(--t1)', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{r.desc}</div>
                    <div style={{ fontSize:11, color:'var(--t3)' }}>{r.date.split('-').reverse().join('.')} · {cat}</div>
                  </div>
                  <div style={{ fontSize:13, fontWeight:600, color:r.type==='income'?'var(--mint)':'var(--coral)', flexShrink:0 }}>
                    {r.type==='income'?'+':'-'}€{r.amount.toLocaleString('de-DE',{minimumFractionDigits:2})}
                  </div>
                </div>
              )
            })}
            {rows.length > 10 && (
              <div style={{ fontSize:12, color:'var(--t3)', textAlign:'center', paddingTop:8 }}>
                + {rows.length - 10} weitere Buchungen
              </div>
            )}
          </Card>

          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={doImport} disabled={importing} style={{ flex:1, justifyContent:'center' }}>
              {importing ? '…Importiere' : `${rows.length} Buchungen importieren`}
            </Btn>
            <Btn variant="ghost" onClick={() => setStep('upload')}>Zurück</Btn>
          </div>
        </div>
      )}

      {step === 'done' && stats && (
        <div style={{ textAlign:'center', padding:'2rem 0' }}>
          <div style={{ fontSize:48, marginBottom:12 }}>✅</div>
          <div style={{ fontSize:18, fontWeight:600, color:'var(--t1)', marginBottom:8 }}>
            Import abgeschlossen!
          </div>
          <div style={{ fontSize:14, color:'var(--t2)', marginBottom:'1.5rem' }}>
            {stats.imported} Buchungen wurden importiert und kategorisiert.
          </div>
          <Btn onClick={onBack} style={{ justifyContent:'center' }}>
            Fertig
          </Btn>
        </div>
      )}
    </div>
  )
}

// ── Erweiterte KI-Regeln für Sparkasse ────────────────────────────────────
// Diese DEFAULT_RULES ersetzen die bisherigen

export const SPARKASSE_RULES = {
  Lebensmittel: 'rewe,lidl,aldi,edeka,penny,netto,kaufland,dm,rossmann,bäcker,metzger,wasgau,backerei,backshop,drogerie,müller,mueller,globus,spar,tegut,real,norma,marktkauf',
  Wohnen:       'miete,nebenkosten,strom,gas,wasser,internet,telefon,hausgeld,1und1,1+1,unitymedia,vodafone,telekom,o2,ewe,stadtwerke,stadtwerk,verbandsgemeinde',
  Transport:    'monatskarte,db,bahn,öpnv,tankstelle,shell,aral,bp,uber,taxi,flixbus,lufthansa,minera,kraftstoff,benzin,parkgebühr,parkgebuehr,autohaus,hyundai,kfz',
  Freizeit:     'netflix,spotify,amazon prime,disney,kino,theater,konzert,steam,playstation,gym,fitnessstudio,fitness,sport,voswinkel,supercell,riot games,epic games,google play,app store,kinoheld,happypottery,pottery,restaurant,pizzeria,ristorante,gastro,diner,eiscafe,sushi,hikari,dong,asia',
  Gesundheit:   'apotheke,arzt,krankenhaus,optiker,zahnarzt,physiotherapie,sanitätshaus,rats-apotheke',
  Versicherung: 'versicherung,allianz,aok,tkk,barmer,huk,continentale,europa verbund,samsung pay,gewerkschaft,verdi',
  Gehalt:       'gehalt,lohn,freelance,honorar,bundeskasse,renten service,rente,pension,lohn  gehalt,rente pensionszahlung,gutschr. ueberw. dauerauftr,zuwendung',
  Sonstiges:    'amazon,topstep,consors,easycredit,teambank,paypal,etsy,microsoft,google,anthropic,claude',
}
