import { useState, useEffect, useCallback } from 'react'
import { db } from './supabase.js'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from 'recharts'
import * as XLSX from 'xlsx'
import './index.css'


// ── Daten & Konstanten ─────────────────────────────────────────────────────

const CATS = ['Wohnen','Lebensmittel','Transport','Freizeit','Gesundheit','Versicherung','Gehalt','Sonstiges']
const SOURCES = ['Bankkonto','Kreditkarte','PayPal','Stripe','Bar']
const FREQS   = ['Monatlich','Wöchentlich','Quartalsweise','Jährlich']

// Pastel-Farben pro Kategorie (light + dark)
const CAT = {
  Wohnen:       { c:'var(--sky)',    bg:'var(--sky-bg)',    hex:'#50A8E0', hexDk:'#70BCEC', icon:'ti-home'          },
  Lebensmittel: { c:'var(--mint)',   bg:'var(--mint-bg)',   hex:'#3EC994', hexDk:'#52C898', icon:'ti-shopping-cart' },
  Transport:    { c:'var(--violet)', bg:'var(--violet-bg)', hex:'#9080D8', hexDk:'#A898E0', icon:'ti-car'           },
  Freizeit:     { c:'var(--rose)',   bg:'var(--rose-bg)',   hex:'#E86880', hexDk:'#F08898', icon:'ti-confetti'      },
  Gesundheit:   { c:'var(--coral)',  bg:'var(--coral-bg)',  hex:'#E87060', hexDk:'#F08878', icon:'ti-heart'         },
  Versicherung: { c:'var(--t2)',     bg:'var(--card2)',     hex:'#6B6A78', hexDk:'#8A8898', icon:'ti-shield'        },
  Gehalt:       { c:'var(--mint)',   bg:'var(--mint-bg)',   hex:'#3EC994', hexDk:'#52C898', icon:'ti-trending-up'   },
  Sonstiges:    { c:'var(--t3)',     bg:'var(--card3)',     hex:'#A4A3B0', hexDk:'#52506A', icon:'ti-dots'          },
}

const SAMPLE = [
  { id:1,  desc:'Gehalt Mai',        amount:2800,  type:'income',  cat:'Gehalt',       date:'2026-05-01', source:'Bankkonto'  },
  { id:2,  desc:'Miete',             amount:850,   type:'expense', cat:'Wohnen',       date:'2026-05-02', source:'Bankkonto'  },
  { id:3,  desc:'REWE',              amount:67.40, type:'expense', cat:'Lebensmittel', date:'2026-05-03', source:'Kreditkarte'},
  { id:4,  desc:'Monatskarte',       amount:49,    type:'expense', cat:'Transport',    date:'2026-05-04', source:'Bankkonto'  },
  { id:5,  desc:'Netflix',           amount:17.99, type:'expense', cat:'Freizeit',     date:'2026-05-05', source:'PayPal'     },
  { id:6,  desc:'Lidl',              amount:43.20, type:'expense', cat:'Lebensmittel', date:'2026-05-07', source:'Kreditkarte'},
  { id:7,  desc:'Kfz-Versicherung',  amount:112,   type:'expense', cat:'Versicherung', date:'2026-05-10', source:'Bankkonto'  },
  { id:8,  desc:'Freelance',         amount:650,   type:'income',  cat:'Gehalt',       date:'2026-05-12', source:'PayPal'     },
  { id:9,  desc:'Apotheke',          amount:28.50, type:'expense', cat:'Gesundheit',   date:'2026-05-14', source:'Kreditkarte'},
  { id:10, desc:'Restaurant',        amount:54,    type:'expense', cat:'Freizeit',     date:'2026-05-16', source:'Kreditkarte'},
  { id:11, desc:'Gehalt April',      amount:2800,  type:'income',  cat:'Gehalt',       date:'2026-04-01', source:'Bankkonto'  },
  { id:12, desc:'Miete April',       amount:850,   type:'expense', cat:'Wohnen',       date:'2026-04-02', source:'Bankkonto'  },
  { id:13, desc:'Einkauf April',     amount:176,   type:'expense', cat:'Lebensmittel', date:'2026-04-10', source:'Kreditkarte'},
  { id:14, desc:'ÖPNV April',        amount:49,    type:'expense', cat:'Transport',     date:'2026-04-04', source:'Bankkonto'  },
  { id:15, desc:'Kino',              amount:28,    type:'expense', cat:'Freizeit',     date:'2026-04-20', source:'Kreditkarte'},
]
const DEF_BUDGETS   = { Wohnen:900, Lebensmittel:400, Transport:100, Freizeit:150, Gesundheit:100, Versicherung:150 }
const DEF_RECURRING = [
  { id:1, desc:'Miete',            amount:850,  type:'expense', cat:'Wohnen',       source:'Bankkonto', freq:'Monatlich', nextDate:'2026-06-01' },
  { id:2, desc:'Netflix',          amount:17.99,type:'expense', cat:'Freizeit',     source:'PayPal',    freq:'Monatlich', nextDate:'2026-06-05' },
  { id:3, desc:'Kfz-Versicherung', amount:112,  type:'expense', cat:'Versicherung', source:'Bankkonto', freq:'Monatlich', nextDate:'2026-06-10' },
]

// ── Helpers ────────────────────────────────────────────────────────────────

const fmt    = n  => n.toLocaleString('de-DE',{ minimumFractionDigits:2, maximumFractionDigits:2 })
const today  = () => new Date().toISOString().split('T')[0]
const MN     = ['','Jan','Feb','Mär','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez']
const fmtMon = m  => MN[parseInt(m.slice(5))] + ' ' + m.slice(2,4)

// ── Hooks ──────────────────────────────────────────────────────────────────
async function syncToSupabase(txs, budgets, recurring) {
  try {
    // Transaktionen sync
  if (txs.length > 0) {
  await db.from('transactions').upsert(
    txs.map(t => ({ id: t.id, description: t.desc, amount: t.amount, type: t.type, cat: t.cat, date: t.date, source: t.source })),
    { onConflict: 'id' }
  )
}
      )
    }
    // Budgets sync
    const budgetRows = Object.entries(budgets).map(([cat, amount]) => ({ cat, amount }))
if (budgetRows.length > 0) await db.from('budgets').upsert(budgetRows, { onConflict: 'cat' })
  } catch(e) { console.log('Sync fehler:', e) }
}

async function loadFromSupabase() {
  try {
    const [txRes, budRes] = await Promise.all([
      db.from('transactions').select('*').order('date', { ascending: false }),
      db.from('budgets').select('*')
    ])
    const txs = (txRes.data || []).map(t => ({ id: t.id, desc: t.description, amount: t.amount, type: t.type, cat: t.cat, date: t.date, source: t.source }))
    const budgets = {}
    ;(budRes.data || []).forEach(b => { budgets[b.cat] = b.amount })
    return { txs, budgets }
  } catch(e) { return null }
}
function useLS(key, def) {
  const [v, setV] = useState(() => {
    try { const s = localStorage.getItem(key); return s ? JSON.parse(s) : def } catch { return def }
  })
  useEffect(() => { try { localStorage.setItem(key, JSON.stringify(v)) } catch {} }, [key, v])
  return [v, setV]
}

function useTheme() {
  const [theme, setTheme] = useLS('ft-theme',
    window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'
  )
  useEffect(() => { document.documentElement.setAttribute('data-theme', theme) }, [theme])
  const toggle = () => setTheme(t => t === 'dark' ? 'light' : 'dark')
  return [theme, toggle]
}

// ── AI-Kategorisierung ─────────────────────────────────────────────────────

const RULES = {
  Lebensmittel: ['rewe','lidl','aldi','edeka','penny','netto','kaufland','dm','rossmann','bäcker','metzger','restaurant','pizza','mcdonalds','burger','kfc','subway','lieferando'],
  Wohnen:       ['miete','nebenkosten','strom','gas','wasser','internet','telefon','hausgeld'],
  Transport:    ['monatskarte','db','bahn','öpnv','tankstelle','shell','aral','bp','uber','taxi','flixbus','lufthansa'],
  Freizeit:     ['netflix','spotify','amazon','disney','kino','theater','konzert','steam','playstation','gym','fitnessstudio'],
  Gesundheit:   ['apotheke','arzt','krankenhaus','optiker','zahnarzt','physiotherapie'],
  Versicherung: ['versicherung','allianz','aok','tkk','barmer','huk'],
  Gehalt:       ['gehalt','lohn','freelance','honorar'],
}

async function aiCat(desc) {
  const low = desc.toLowerCase()
  for (const [cat, keywords] of Object.entries(RULES)) {
    if (keywords.some(k => low.includes(k))) return cat
  }
  return 'Sonstiges'
}


// ── Shared UI-Komponenten ──────────────────────────────────────────────────

function Card({ children, style = {}, onClick }) {
  return (
    <div onClick={onClick} style={{
      background: 'var(--card)', borderRadius: 'var(--r-lg)',
      border: '1px solid var(--brd)', boxShadow: 'var(--sh-sm)',
      padding: '1.1rem 1.25rem', marginBottom: '0.75rem',
      transition: 'box-shadow 0.2s', cursor: onClick ? 'pointer' : 'default',
      ...style
    }}>{children}</div>
  )
}

function Pill({ children, color, bg }) {
  return (
    <span style={{
      display: 'inline-flex', alignItems: 'center', gap: 4,
      background: bg, color: color, fontSize: 11, fontWeight: 500,
      padding: '3px 9px', borderRadius: 100, whiteSpace: 'nowrap'
    }}>{children}</span>
  )
}

function CatIcon({ cat, size = 34 }) {
  const info = CAT[cat] || CAT.Sonstiges
  return (
    <div style={{
      width: size, height: size, borderRadius: '50%',
      background: info.bg, color: info.c, flexShrink: 0,
      display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: size * 0.44
    }}>
      <i className={`ti ${info.icon}`} aria-hidden="true" />
    </div>
  )
}

function Btn({ children, variant = 'primary', onClick, disabled, style = {} }) {
  const base = {
    display: 'inline-flex', alignItems: 'center', gap: 6,
    padding: '9px 18px', borderRadius: 'var(--r-md)', fontSize: 14,
    fontWeight: 500, cursor: disabled ? 'not-allowed' : 'pointer',
    opacity: disabled ? 0.6 : 1, transition: 'opacity 0.15s, transform 0.1s',
    border: '1px solid transparent', ...style
  }
  const v = {
    primary: { background: 'var(--t1)', color: 'var(--page)' },
    ghost:   { background: 'var(--card2)', color: 'var(--t1)', border: '1px solid var(--brd)' },
    danger:  { background: 'var(--coral-bg)', color: 'var(--coral)' },
  }
  return <button style={{ ...base, ...v[variant] }} onClick={onClick} disabled={disabled}>{children}</button>
}

function Field({ label, children }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
      <label style={{ fontSize: 12, fontWeight: 500, color: 'var(--t2)' }}>{label}</label>
      {children}
    </div>
  )
}

const inp = {
  width: '100%', height: 40, padding: '0 12px',
  borderRadius: 'var(--r-md)', border: '1px solid var(--brd2)',
  background: 'var(--card)', color: 'var(--t1)', fontSize: 14,
}

function Input({ style, ...props }) {
  return <input style={{ ...inp, ...style }} {...props} />
}

function Select({ children, style, ...props }) {
  return (
    <select style={{ ...inp, ...style, cursor: 'pointer' }} {...props}>
      {children}
    </select>
  )
}

// ── Haupt-App ──────────────────────────────────────────────────────────────

export default function App() {
  const [tab, setTab]           = useState('dashboard')
  const [txs, setTxs]           = useLS('ft-txs', SAMPLE)
  const [budgets, setBudgets]   = useLS('ft-budgets', DEF_BUDGETS)
  const [recurring, setRec]     = useLS('ft-recurring', DEF_RECURRING)
  const [month, setMonth]       = useLS('ft-month', '2026-05')
  const [nid, setNid]           = useLS('ft-nid', 200)
// Supabase laden beim Start
useEffect(() => {
  loadFromSupabase().then(data => {
    if (data?.txs?.length > 0) setTxs(data.txs)
    if (Object.keys(data?.budgets || {}).length > 0) setBudgets(data.budgets)
  })
}, [])

// Supabase sync bei jeder Änderung
useEffect(() => {
  syncToSupabase(txs, budgets, recurring)
}, [txs, budgets, recurring])
  const [theme, toggleTheme]    = useTheme()
  const [toast, setToast]       = useState(null)

  const showToast = useCallback((msg, type = 'ok') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 2600)
  }, [])

  const months = [...new Set(txs.map(t => t.date.slice(0,7)))].sort().reverse()
  const filtered = txs.filter(t => month === 'all' || t.date.startsWith(month))
  const income  = filtered.filter(t => t.type === 'income').reduce((s,t) => s + t.amount, 0)
  const expense = filtered.filter(t => t.type === 'expense').reduce((s,t) => s + t.amount, 0)
  const catSpend = {}
  filtered.filter(t => t.type === 'expense').forEach(t => { catSpend[t.cat] = (catSpend[t.cat]||0) + t.amount })

  const warnings = Object.entries(budgets)
    .filter(([c,lim]) => lim > 0 && (catSpend[c]||0) >= lim * 0.8)

  const NAV = [
    { id:'dashboard',    icon:'ti-layout-dashboard', label:'Übersicht' },
    { id:'transactions', icon:'ti-list-details',     label:'Buchungen' },
    { id:'budget',       icon:'ti-target',           label:'Budget'    },
    { id:'recurring',    icon:'ti-repeat',           label:'Fixkosten' },
    { id:'more',         icon:'ti-dots',             label:'Mehr'      },
  ]

  return (
    <div style={{ display:'flex', flexDirection:'column', height:'100%', background:'var(--page)' }}>

      {/* ── Header ── */}
      <header style={{
        position:'sticky', top:0, zIndex:50,
        background:'var(--card)', borderBottom:'1px solid var(--brd)',
        backdropFilter:'blur(12px)', WebkitBackdropFilter:'blur(12px)',
        padding:'0 1rem', paddingTop:'env(safe-area-inset-top,0)',
        display:'flex', alignItems:'center', gap:10, minHeight:60
      }}>
        <div style={{ fontSize:18, fontWeight:600, color:'var(--t1)', flex:1, letterSpacing:'-0.3px' }}>
          💰 Finanzen
        </div>

        <Select value={month} onChange={e => setMonth(e.target.value)}
          style={{ width:'auto', height:34, fontSize:13, paddingRight:28 }}>
          <option value="all">Alle Monate</option>
          {months.map(m => <option key={m} value={m}>{fmtMon(m)}</option>)}
        </Select>

        <button onClick={toggleTheme} style={{
          width:34, height:34, borderRadius:'var(--r-md)',
          background:'var(--card2)', border:'1px solid var(--brd)',
          display:'flex', alignItems:'center', justifyContent:'center',
          color:'var(--t2)', fontSize:16, flexShrink:0
        }} aria-label="Design wechseln">
          <i className={`ti ${theme === 'dark' ? 'ti-sun' : 'ti-moon'}`} aria-hidden="true" />
        </button>
      </header>

      {/* ── Budget-Warnungen ── */}
      {warnings.length > 0 && (
        <div style={{ background:'var(--honey-bg)', borderBottom:'1px solid var(--honey-dim)', padding:'10px 1rem' }}>
          {warnings.map(([cat, lim]) => {
            const spent = catSpend[cat]||0, pct = Math.round(spent/lim*100), over = spent > lim
            return (
              <div key={cat} style={{ fontSize:13, color:'var(--honey)', display:'flex', alignItems:'center', gap:6, marginBottom:2 }}>
                <i className="ti ti-alert-triangle" aria-hidden="true" />
                {over
                  ? `Budget überschritten: ${cat} · €${fmt(spent)} / €${fmt(lim)}`
                  : `Budget fast voll: ${cat} · ${pct}%`}
              </div>
            )
          })}
        </div>
      )}

      {/* ── Scrollbarer Inhalt ── */}
      <main style={{ flex:1, overflowY:'auto', paddingBottom:`calc(var(--nav-h) + 8px)` }}>
        <div style={{ padding:'1rem', maxWidth:680, margin:'0 auto' }}>
          {tab === 'dashboard'    && <Dashboard filtered={filtered} income={income} expense={expense}
                                      catSpend={catSpend} allTxs={txs} theme={theme} />}
          {tab === 'transactions' && <Transactions txs={filtered} setTxs={setTxs}
                                      nid={nid} setNid={setNid} toast={showToast} theme={theme} />}
          {tab === 'budget'       && <Budget budgets={budgets} setBudgets={setBudgets}
                                      catSpend={catSpend} toast={showToast} />}
          {tab === 'recurring'    && <Recurring recurring={recurring} setRec={setRec}
                                      setTxs={setTxs} nid={nid} setNid={setNid} toast={showToast} />}
          {tab === 'more'         && <More txs={filtered} month={month} income={income}
                                      expense={expense} setTxs={setTxs} nid={nid} setNid={setNid} toast={showToast} />}
        </div>
      </main>

      {/* ── Bottom Nav ── */}
      <nav style={{
        position:'fixed', bottom:0, left:0, right:0, zIndex:50,
        height:`calc(var(--nav-h) + env(safe-area-inset-bottom,0px))`,
        background:'var(--card)', borderTop:'1px solid var(--brd)',
        backdropFilter:'blur(16px)', WebkitBackdropFilter:'blur(16px)',
        display:'flex', alignItems:'flex-start', paddingTop:4,
        paddingBottom:'env(safe-area-inset-bottom,0px)',
      }}>
        {NAV.map(n => {
          const active = tab === n.id
          return (
            <button key={n.id} onClick={() => setTab(n.id)} style={{
              flex:1, display:'flex', flexDirection:'column', alignItems:'center',
              gap:3, padding:'6px 0', border:'none', background:'none', cursor:'pointer',
              color: active ? 'var(--violet)' : 'var(--t3)', transition:'color 0.2s',
            }}>
              <div style={{
                width:38, height:28, borderRadius:14, display:'flex', alignItems:'center',
                justifyContent:'center', fontSize:18,
                background: active ? 'var(--violet-bg)' : 'transparent',
                transition:'background 0.2s',
              }}>
                <i className={`ti ${n.icon}`} aria-hidden="true" />
              </div>
              <span style={{ fontSize:10, fontWeight: active ? 600 : 400 }}>{n.label}</span>
            </button>
          )
        })}
      </nav>

      {/* ── Toast ── */}
      {toast && (
        <div style={{
          position:'fixed', bottom:`calc(var(--nav-h) + 12px)`, left:'50%',
          transform:'translateX(-50%)', zIndex:200,
          background: toast.type === 'err' ? 'var(--coral)' : 'var(--t1)',
          color:'var(--page)', padding:'10px 20px', borderRadius:'var(--r-xl)',
          fontSize:13, fontWeight:500, boxShadow:'var(--sh-lg)',
          animation:'fadeUp 0.2s ease', whiteSpace:'nowrap',
        }}>
          {toast.msg}
        </div>
      )}
    </div>
  )
}

// ── Dashboard ──────────────────────────────────────────────────────────────

function Dashboard({ filtered, income, expense, catSpend, allTxs, theme }) {
  const balance  = income - expense
  const dk       = theme === 'dark'

  const catData = Object.entries(catSpend)
    .map(([name, value]) => ({ name, value: Math.round(value * 100) / 100 }))
    .sort((a,b) => b.value - a.value)

  const monthMap = {}
  allTxs.forEach(t => {
    const m = t.date.slice(0,7)
    if (!monthMap[m]) monthMap[m] = { name: fmtMon(m), inc:0, exp:0 }
    if (t.type === 'income') monthMap[m].inc += t.amount
    else monthMap[m].exp += t.amount
  })
  const monthData = Object.values(monthMap)
    .sort((a,b) => a.name.localeCompare(b.name))
    .map(d => ({ ...d, inc: Math.round(d.inc), exp: Math.round(d.exp) }))

  const recentTxs = [...filtered].sort((a,b) => b.date.localeCompare(a.date)).slice(0,5)

  const metrics = [
    { label:'Einnahmen',  val:`+€${fmt(income)}`,                                        c:'var(--mint)',   bg:'var(--mint-bg)'   },
    { label:'Ausgaben',   val:`-€${fmt(expense)}`,                                        c:'var(--coral)',  bg:'var(--coral-bg)'  },
    { label:'Saldo',      val:`${balance>=0?'+':''}€${fmt(balance)}`,                     c:'var(--violet)', bg:'var(--violet-bg)' },
    { label:'Buchungen',  val:filtered.length,                                             c:'var(--sky)',    bg:'var(--sky-bg)'    },
  ]

  return (
    <div className="fade-up">
      {/* Metric Cards */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'0.6rem', marginBottom:'1.25rem' }}>
        {metrics.map(m => (
          <div key={m.label} style={{
            background: m.bg, borderRadius:'var(--r-lg)',
            padding:'1rem 1.1rem', border:'1px solid transparent'
          }}>
            <div style={{ fontSize:12, fontWeight:500, color: m.c, marginBottom:6, opacity:0.8 }}>{m.label}</div>
            <div style={{ fontSize:22, fontWeight:600, color: m.c, letterSpacing:'-0.5px', fontVariantNumeric:'tabular-nums' }}>
              {m.val}
            </div>
          </div>
        ))}
      </div>

      {/* Kategorie-Chart */}
      <Card>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:'1rem' }}>
          Ausgaben nach Kategorie
        </div>
        {catData.length === 0
          ? <EmptyState icon="ti-chart-bar" text="Keine Ausgaben im Zeitraum" />
          : <ResponsiveContainer width="100%" height={180}>
              <BarChart data={catData} margin={{ top:0, right:0, left:-20, bottom:0 }}>
                <XAxis dataKey="name" tick={{ fontSize:10, fill:'var(--t3)' }} />
                <YAxis tick={{ fontSize:10, fill:'var(--t3)' }} tickFormatter={v=>'€'+v} />
                <Tooltip
                  formatter={v => [`€${fmt(v)}`, 'Ausgaben']}
                  contentStyle={{ fontSize:12, borderRadius:10, border:'none', background:'var(--card)', color:'var(--t1)', boxShadow:'var(--sh-md)' }}
                  cursor={{ fill:'var(--card2)' }}
                />
                <Bar dataKey="value" radius={[6,6,0,0]}>
                  {catData.map(e => (
                    <Cell key={e.name} fill={dk ? (CAT[e.name]?.hexDk||'#8A8898') : (CAT[e.name]?.hex||'#A4A3B0')} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
        }
      </Card>

      {/* Monatsverlauf */}
      <Card>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:'0.6rem' }}>Monatsverlauf</div>
        <div style={{ display:'flex', gap:16, marginBottom:'0.75rem' }}>
          {[{ col: dk?'#52C898':'#3EC994', label:'Einnahmen' }, { col: dk?'#F08878':'#E87060', label:'Ausgaben' }].map(l => (
            <span key={l.label} style={{ display:'flex', alignItems:'center', gap:5, fontSize:12, color:'var(--t2)' }}>
              <span style={{ width:10, height:10, borderRadius:3, background:l.col, flexShrink:0 }} />
              {l.label}
            </span>
          ))}
        </div>
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={monthData} margin={{ top:0, right:0, left:-20, bottom:0 }}>
            <XAxis dataKey="name" tick={{ fontSize:10, fill:'var(--t3)' }} />
            <YAxis tick={{ fontSize:10, fill:'var(--t3)' }} tickFormatter={v=>'€'+v} />
            <Tooltip
              formatter={(v,n) => [`€${fmt(v)}`, n==='inc'?'Einnahmen':'Ausgaben']}
              contentStyle={{ fontSize:12, borderRadius:10, border:'none', background:'var(--card)', color:'var(--t1)', boxShadow:'var(--sh-md)' }}
              cursor={{ fill:'var(--card2)' }}
            />
            <Bar dataKey="inc" fill={dk?'#52C898':'#3EC994'} radius={[4,4,0,0]} />
            <Bar dataKey="exp" fill={dk?'#F08878':'#E87060'} radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>

      {/* Letzte Buchungen */}
      <Card>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:'0.75rem' }}>Letzte Buchungen</div>
        {recentTxs.length === 0
          ? <EmptyState icon="ti-receipt" text="Keine Buchungen" />
          : recentTxs.map((t, i) => (
              <div key={t.id} style={{
                display:'flex', alignItems:'center', gap:10, padding:'8px 0',
                borderTop: i > 0 ? '1px solid var(--brd)' : 'none'
              }}>
                <CatIcon cat={t.cat} size={32} />
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{t.desc}</div>
                  <div style={{ fontSize:11, color:'var(--t3)' }}>{t.date.split('-').reverse().join('.')}</div>
                </div>
                <div style={{ fontSize:14, fontWeight:600, color: t.type==='income'?'var(--mint)':'var(--coral)', flexShrink:0, fontVariantNumeric:'tabular-nums' }}>
                  {t.type==='income'?'+':'-'}€{fmt(t.amount)}
                </div>
              </div>
            ))
        }
      </Card>
    </div>
  )
}

// ── Buchungen ──────────────────────────────────────────────────────────────

function Transactions({ txs, setTxs, nid, setNid, toast, theme }) {
  const BLANK = { desc:'', amount:'', type:'expense', cat:'Sonstiges', date:today(), source:'Bankkonto' }
  const [form, setForm]       = useState(null)
  const [aiLoad, setAiLoad]   = useState(false)
  const [typeF, setTypeF]     = useState('all')
  const [catF, setCatF]       = useState('all')
  const [q, setQ]             = useState('')

  const set = k => e => setForm(p => ({ ...p, [k]: e.target.value }))

  const sorted = [...txs]
    .filter(t => (typeF==='all'||t.type===typeF) && (catF==='all'||t.cat===catF) && (!q||t.desc.toLowerCase().includes(q.toLowerCase())))
    .sort((a,b) => b.date.localeCompare(a.date))

  const handleAI = async () => {
    if (!form?.desc) { toast('Beschreibung eingeben', 'err'); return }
    setAiLoad(true); toast('KI analysiert…')
    const cat = await aiCat(form.desc)
    setForm(p => ({ ...p, cat })); toast('Kategorie: ' + cat)
    setAiLoad(false)
  }

  const save = () => {
    if (!form.desc||!form.amount||!form.date) { toast('Pflichtfelder ausfüllen', 'err'); return }
setTxs(p => [{ ...form, id:nid, amount:parseFloat(form.amount) }, ...p])
    setNid(n => n+1); setForm(null); toast('Buchung gespeichert ✓')
  }

  const del = id => { setTxs(p => p.filter(t => t.id !== id)); toast('Gelöscht') }

  return (
    <div className="fade-up">
      {/* Filter + Add */}
      <div style={{ display:'flex', gap:8, marginBottom:'0.75rem', flexWrap:'wrap' }}>
        <Select value={typeF} onChange={e=>setTypeF(e.target.value)} style={{ flex:1, minWidth:100, height:38 }}>
          <option value="all">Alle Typen</option>
          <option value="income">Einnahmen</option>
          <option value="expense">Ausgaben</option>
        </Select>
        <Select value={catF} onChange={e=>setCatF(e.target.value)} style={{ flex:1, minWidth:120, height:38 }}>
          <option value="all">Alle Kategorien</option>
          {CATS.map(c => <option key={c}>{c}</option>)}
        </Select>
      </div>
      <div style={{ display:'flex', gap:8, marginBottom:'1rem' }}>
        <Input value={q} onChange={e=>setQ(e.target.value)} placeholder="Suchen…" style={{ flex:1, height:38 }} />
        <Btn onClick={() => setForm(BLANK)}>
          <i className="ti ti-plus" aria-hidden="true" /> Neu
        </Btn>
      </div>

      {/* Formular */}
      {form && (
        <div style={{ background:'var(--card2)', borderRadius:'var(--r-lg)', padding:'1.1rem', marginBottom:'1rem', border:'1px solid var(--brd)' }}>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:10 }}>
            <Field label="Beschreibung" style={{ gridColumn:'1/-1' }}>
              <div style={{ display:'flex', gap:6, gridColumn:'1/-1' }}>
                <Input value={form.desc} onChange={set('desc')} placeholder="z.B. REWE, Miete…" style={{ flex:1 }} />
                <button onClick={handleAI} disabled={aiLoad} title="KI-Kategorisierung" style={{
                  height:40, padding:'0 12px', borderRadius:'var(--r-md)',
                  background:'var(--violet-bg)', color:'var(--violet)',
                  border:'1px solid var(--violet-dim)', fontSize:13, fontWeight:500,
                  cursor: aiLoad ? 'wait' : 'pointer', whiteSpace:'nowrap', flexShrink:0
                }}>
                  {aiLoad ? '…' : '🤖 KI'}
                </button>
              </div>
            </Field>
          </div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:10 }}>
            <Field label="Betrag (€)">
              <Input type="number" value={form.amount} onChange={set('amount')} placeholder="0.00" step="0.01" />
            </Field>
            <Field label="Typ">
              <Select value={form.type} onChange={set('type')}>
                <option value="expense">Ausgabe</option>
                <option value="income">Einnahme</option>
              </Select>
            </Field>
            <Field label="Kategorie">
              <Select value={form.cat} onChange={set('cat')}>
                {CATS.map(c => <option key={c}>{c}</option>)}
              </Select>
            </Field>
            <Field label="Datum">
              <Input type="date" value={form.date} onChange={set('date')} />
            </Field>
            <Field label="Quelle">
              <Select value={form.source} onChange={set('source')}>
                {SOURCES.map(s => <option key={s}>{s}</option>)}
              </Select>
            </Field>
          </div>
          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={save}>Speichern</Btn>
            <Btn variant="ghost" onClick={() => setForm(null)}>Abbrechen</Btn>
          </div>
        </div>
      )}

      {/* Liste */}
      {sorted.length === 0
        ? <EmptyState icon="ti-receipt" text="Keine Buchungen gefunden" />
        : sorted.map((t, i) => (
            <div key={t.id} style={{
              background:'var(--card)', borderRadius:'var(--r-lg)', marginBottom:6,
              border:'1px solid var(--brd)', display:'flex', alignItems:'center', gap:10,
              padding:'10px 12px', boxShadow:'var(--sh-sm)'
            }}>
              <CatIcon cat={t.cat} size={36} />
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{t.desc}</div>
                <div style={{ display:'flex', alignItems:'center', gap:6, marginTop:2, flexWrap:'wrap' }}>
                  <span style={{ fontSize:11, color:'var(--t3)' }}>{t.date.split('-').reverse().join('.')}</span>
                  <Pill color={CAT[t.cat]?.c||'var(--t3)'} bg={CAT[t.cat]?.bg||'var(--card2)'}>{t.cat}</Pill>
                  <span style={{ fontSize:11, color:'var(--t3)' }}>{t.source}</span>
                </div>
              </div>
              <div style={{ fontSize:15, fontWeight:600, color: t.type==='income'?'var(--mint)':'var(--coral)', flexShrink:0, fontVariantNumeric:'tabular-nums' }}>
                {t.type==='income'?'+':'-'}€{fmt(t.amount)}
              </div>
              <button onClick={() => del(t.id)} style={{
                width:28, height:28, borderRadius:'var(--r-sm)', display:'flex', alignItems:'center',
                justifyContent:'center', color:'var(--t3)', fontSize:14, flexShrink:0,
                background:'transparent', border:'none', cursor:'pointer',
              }} aria-label="Löschen">
                <i className="ti ti-x" aria-hidden="true" />
              </button>
            </div>
          ))
      }
    </div>
  )
}

// ── Budget ─────────────────────────────────────────────────────────────────

function Budget({ budgets, setBudgets, catSpend, toast }) {
  const [editing, setEditing] = useState(null)
  const [editVal, setEditVal] = useState('')

  const expCats = CATS.filter(c => c !== 'Gehalt')
  const totalBudget = expCats.reduce((s,c) => s + (budgets[c]||0), 0)
  const totalSpent  = expCats.reduce((s,c) => s + (catSpend[c]||0), 0)

  const saveEdit = cat => {
    const v = parseFloat(editVal)
    if (isNaN(v) || v < 0) { toast('Ungültiger Betrag', 'err'); return }
    setBudgets(b => ({ ...b, [cat]: v }))
    setEditing(null); toast('Budget gespeichert ✓')
  }

  return (
    <div className="fade-up">
      {/* Gesamt */}
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:'0.6rem', marginBottom:'1.25rem' }}>
        {[
          { label:'Gesamtbudget', val:`€${fmt(totalBudget)}`, c:'var(--violet)', bg:'var(--violet-bg)' },
          { label:'Ausgegeben',   val:`€${fmt(totalSpent)}`,  c: totalSpent>totalBudget?'var(--coral)':'var(--t1)', bg:'var(--card2)' },
          { label:'Verbleibend',  val:`€${fmt(Math.max(totalBudget-totalSpent,0))}`, c:'var(--mint)', bg:'var(--mint-bg)' },
        ].map(m => (
          <div key={m.label} style={{ background:m.bg, borderRadius:'var(--r-lg)', padding:'0.85rem 1rem' }}>
            <div style={{ fontSize:11, fontWeight:500, color:m.c, opacity:0.8, marginBottom:4 }}>{m.label}</div>
            <div style={{ fontSize:17, fontWeight:600, color:m.c, fontVariantNumeric:'tabular-nums' }}>{m.val}</div>
          </div>
        ))}
      </div>

      {expCats.map(cat => {
        const lim   = budgets[cat] || 0
        const spent = catSpend[cat] || 0
        const pct   = lim > 0 ? Math.min(spent / lim * 100, 100) : 0
        const over  = spent > lim && lim > 0
        const warn  = pct >= 80 && !over
        const barC  = over ? 'var(--coral)' : warn ? 'var(--honey)' : 'var(--mint)'
        const info  = CAT[cat] || CAT.Sonstiges

        return (
          <Card key={cat} style={{ padding:'1rem 1.1rem' }}>
            <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:10, flexWrap:'wrap' }}>
              <CatIcon cat={cat} size={32} />
              <div style={{ flex:1 }}>
                <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                  <span style={{ fontSize:14, fontWeight:600, color:'var(--t1)' }}>{cat}</span>
                  {over && <Pill color="var(--coral)" bg="var(--coral-bg)">Überschritten</Pill>}
                  {warn && <Pill color="var(--honey)" bg="var(--honey-bg)">Fast voll</Pill>}
                </div>
              </div>
              <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                <span style={{ fontSize:13, color:'var(--t2)', fontVariantNumeric:'tabular-nums' }}>€{fmt(spent)} /</span>
                {editing === cat ? (
                  <>
                    <input type="number" value={editVal} onChange={e=>setEditVal(e.target.value)}
                      onKeyDown={e=>e.key==='Enter'&&saveEdit(cat)}
                      autoFocus style={{ ...inp, width:80, height:32, fontSize:13 }} />
                    <button onClick={()=>saveEdit(cat)} style={{ ...inp, width:'auto', height:32, padding:'0 10px', background:'var(--mint)', color:'#fff', border:'none', fontWeight:500, fontSize:13, cursor:'pointer', borderRadius:'var(--r-sm)' }}>OK</button>
                  </>
                ) : (
                  <>
                    <span style={{ fontSize:13, fontWeight:600, color:'var(--t1)', fontVariantNumeric:'tabular-nums' }}>€{fmt(lim)}</span>
                    <button onClick={()=>{setEditing(cat);setEditVal(String(lim))}} style={{ color:'var(--t3)', fontSize:14, padding:2 }} aria-label="Bearbeiten">
                      <i className="ti ti-edit" aria-hidden="true" />
                    </button>
                  </>
                )}
              </div>
            </div>
            <div style={{ height:8, background:'var(--card2)', borderRadius:100, overflow:'hidden' }}>
              <div style={{ height:'100%', width:`${pct}%`, background:barC, borderRadius:100, transition:'width 0.4s ease' }} />
            </div>
            <div style={{ fontSize:11, color:'var(--t3)', marginTop:5 }}>
              {Math.round(pct)}% verwendet{lim>0?` · €${fmt(Math.max(lim-spent,0))} verbleibend`:''}
            </div>
          </Card>
        )
      })}
    </div>
  )
}

// ── Wiederkehrend ──────────────────────────────────────────────────────────

function Recurring({ recurring, setRec, setTxs, nid, setNid, toast }) {
  const BLANK = { desc:'', amount:'', type:'expense', cat:'Sonstiges', source:'Bankkonto', freq:'Monatlich', nextDate:'' }
  const [form, setForm] = useState(null)
  const [rNid, setRNid] = useState(500)
  const set = k => e => setForm(p => ({ ...p, [k]: e.target.value }))

  const monthlyFixed = recurring.filter(r=>r.freq==='Monatlich'&&r.type==='expense').reduce((s,r)=>s+r.amount,0)

  const save = () => {
    if (!form.desc||!form.amount||!form.nextDate) { toast('Pflichtfelder ausfüllen','err'); return }
    setRec(r => [...r, { ...form, id:rNid, amount:parseFloat(form.amount) }])
    setRNid(n=>n+1); setForm(null); toast('Fixkosten gespeichert ✓')
  }
  const del = id => { setRec(r=>r.filter(x=>x.id!==id)); toast('Gelöscht') }
  const applyAll = () => {
    const t0 = today(); let id = nid
    const newTxs = recurring.map(r => ({ id:id++, desc:r.desc, amount:r.amount, type:r.type, cat:r.cat, date:t0, source:r.source }))
    setTxs(p=>[...p,...newTxs]); setNid(id)
    toast(`${recurring.length} Buchungen eingetragen ✓`)
  }

  return (
    <div className="fade-up">
      {recurring.length > 0 && (
        <div style={{ background:'var(--coral-bg)', borderRadius:'var(--r-lg)', padding:'1rem 1.1rem', marginBottom:'1rem', display:'flex', alignItems:'center', justifyContent:'space-between', gap:8, flexWrap:'wrap' }}>
          <div>
            <div style={{ fontSize:12, color:'var(--coral)', fontWeight:500, marginBottom:3 }}>Feste monatliche Ausgaben</div>
            <div style={{ fontSize:22, fontWeight:600, color:'var(--coral)', fontVariantNumeric:'tabular-nums' }}>-€{fmt(monthlyFixed)}</div>
          </div>
          <Btn variant="ghost" onClick={applyAll}>
            <i className="ti ti-player-play" aria-hidden="true" /> Heute eintragen
          </Btn>
        </div>
      )}

      <div style={{ marginBottom:'1rem' }}>
        <Btn onClick={()=>setForm(BLANK)}>
          <i className="ti ti-plus" aria-hidden="true" /> Neue Fixkosten
        </Btn>
      </div>

      {form && (
        <div style={{ background:'var(--card2)', borderRadius:'var(--r-lg)', padding:'1.1rem', marginBottom:'1rem', border:'1px solid var(--brd)' }}>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:10 }}>
            <Field label="Beschreibung"><Input value={form.desc} onChange={set('desc')} placeholder="z.B. Miete" /></Field>
            <Field label="Betrag (€)"><Input type="number" value={form.amount} onChange={set('amount')} step="0.01" /></Field>
            <Field label="Typ"><Select value={form.type} onChange={set('type')}><option value="expense">Ausgabe</option><option value="income">Einnahme</option></Select></Field>
            <Field label="Kategorie"><Select value={form.cat} onChange={set('cat')}>{CATS.map(c=><option key={c}>{c}</option>)}</Select></Field>
            <Field label="Frequenz"><Select value={form.freq} onChange={set('freq')}>{FREQS.map(f=><option key={f}>{f}</option>)}</Select></Field>
            <Field label="Nächstes Datum"><Input type="date" value={form.nextDate} onChange={set('nextDate')} /></Field>
            <Field label="Quelle"><Select value={form.source} onChange={set('source')}>{SOURCES.map(s=><option key={s}>{s}</option>)}</Select></Field>
          </div>
          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={save}>Speichern</Btn>
            <Btn variant="ghost" onClick={()=>setForm(null)}>Abbrechen</Btn>
          </div>
        </div>
      )}

      {recurring.length === 0 && !form
        ? <EmptyState icon="ti-repeat" text="Noch keine Fixkosten definiert" />
        : recurring.map(r => (
            <Card key={r.id} style={{ display:'flex', alignItems:'center', gap:10 }}>
              <CatIcon cat={r.cat} size={36} />
              <div style={{ flex:1, minWidth:0 }}>
                <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)' }}>{r.desc}</div>
                <div style={{ fontSize:11, color:'var(--t3)', marginTop:2 }}>
                  {r.freq} · {r.cat} · Nächstes: {r.nextDate?.split('-').reverse().join('.')}
                </div>
              </div>
              <div style={{ fontSize:15, fontWeight:600, color:r.type==='income'?'var(--mint)':'var(--coral)', fontVariantNumeric:'tabular-nums', flexShrink:0 }}>
                {r.type==='income'?'+':'-'}€{fmt(r.amount)}
              </div>
              <button onClick={()=>del(r.id)} style={{ color:'var(--t3)', fontSize:15, padding:4, flexShrink:0 }} aria-label="Löschen">
                <i className="ti ti-trash" aria-hidden="true" />
              </button>
            </Card>
          ))
      }
    </div>
  )
}

// ── Mehr (Import + Export) ─────────────────────────────────────────────────

function More({ txs, month, income, expense, setTxs, nid, setNid, toast }) {
  const [view, setView] = useState('menu')

  if (view === 'import') return <ImportView setTxs={setTxs} nid={nid} setNid={setNid} toast={toast} onBack={()=>setView('menu')} />
  if (view === 'export') return <ExportView txs={txs} month={month} income={income} expense={expense} toast={toast} onBack={()=>setView('menu')} />

  const items = [
    { id:'import', icon:'ti-upload', label:'CSV importieren', sub:'Kontoauszug von Bank, PayPal oder Stripe hochladen', c:'var(--sky)', bg:'var(--sky-bg)' },
    { id:'export', icon:'ti-download', label:'Exportieren', sub:'Excel, CSV oder druckbaren PDF-Bericht erstellen', c:'var(--mint)', bg:'var(--mint-bg)' },
  ]

  return (
    <div className="fade-up">
      <div style={{ fontSize:18, fontWeight:600, color:'var(--t1)', marginBottom:'1.25rem' }}>Mehr</div>
      {items.map(it => (
        <Card key={it.id} onClick={()=>setView(it.id)} style={{ display:'flex', alignItems:'center', gap:14, cursor:'pointer' }}>
          <div style={{ width:44, height:44, borderRadius:'var(--r-md)', background:it.bg, color:it.c, display:'flex', alignItems:'center', justifyContent:'center', fontSize:22, flexShrink:0 }}>
            <i className={`ti ${it.icon}`} aria-hidden="true" />
          </div>
          <div style={{ flex:1 }}>
            <div style={{ fontSize:15, fontWeight:600, color:'var(--t1)', marginBottom:2 }}>{it.label}</div>
            <div style={{ fontSize:12, color:'var(--t2)', lineHeight:1.4 }}>{it.sub}</div>
          </div>
          <i className="ti ti-chevron-right" style={{ color:'var(--t3)', fontSize:16, flexShrink:0 }} aria-hidden="true" />
        </Card>
      ))}

      <div style={{ marginTop:'1.5rem', background:'var(--violet-bg)', borderRadius:'var(--r-lg)', padding:'1rem 1.1rem' }}>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--violet)', marginBottom:6 }}>
          <i className="ti ti-building-bank" style={{ marginRight:6 }} aria-hidden="true" />
          Bank automatisch verbinden
        </div>
        <div style={{ fontSize:12, color:'var(--t2)', lineHeight:1.6 }}>
          Echtzeit-Bankanbindung über GoCardless Open Banking (PSD2).
          Unterstützt alle deutschen Banken — DKB, ING, Sparkasse, Volksbank u.v.m.
          <br /><br />
          <strong style={{ color:'var(--t1)' }}>Setup in README.md</strong> — kostenlos für Privatpersonen.
        </div>
      </div>
    </div>
  )
}

// ── Import-View ────────────────────────────────────────────────────────────

function ImportView({ setTxs, nid, setNid, toast, onBack }) {
  const [headers, setHeaders]   = useState([])
  const [rows, setRows]         = useState([])
  const [preview, setPreview]   = useState([])
  const [map, setMap]           = useState({ date:'', desc:'', amount:'' })
  const [importing, setImport]  = useState(false)

  const handleFile = e => {
    const file = e.target.files[0]; if (!file) return
    const reader = new FileReader()
    reader.onload = ev => {
      const text = ev.target.result
      const lines = text.split('\n').filter(l => l.trim())
      const sep = lines[0].includes(';') ? ';' : ','
      const h = lines[0].split(sep).map(x => x.trim().replace(/^"|"$/g,''))
      const r = lines.slice(1).map(l => l.split(sep).map(x => x.trim().replace(/^"|"$/g,'')))
      setHeaders(h); setRows(r); setPreview(r.slice(0,4))
      setMap({
        date:   h.find(x=>/datum|date|buchungstag/i.test(x))||'',
        desc:   h.find(x=>/beschreibung|verwendung|name|payee|memo/i.test(x))||'',
        amount: h.find(x=>/betrag|amount|summe|umsatz/i.test(x))||'',
      })
      toast(`${r.length} Zeilen erkannt`)
    }
    reader.readAsText(file, 'UTF-8')
  }

  const doImport = async () => {
    if (!map.date||!map.desc||!map.amount) { toast('Alle Spalten zuordnen', 'err'); return }
    setImport(true)
    const di=headers.indexOf(map.date), ni=headers.indexOf(map.desc), ai=headers.indexOf(map.amount)
    let id=nid; const newTxs=[]
    for (let i=0; i<rows.length; i++) {
      const row=rows[i]
      const rawAmt=parseFloat(row[ai]?.replace(',','.').replace(/[^0-9.-]/g,''))
      if (isNaN(rawAmt)) continue
      let date=row[di]||today()
      if (/^\d{2}\.\d{2}\.\d{4}$/.test(date)) { const [d,m,y]=date.split('.'); date=`${y}-${m}-${d}` }
      let cat='Sonstiges'
      if (i<8) { try { cat=await aiCat(row[ni]||'') } catch {} }
      newTxs.push({ id:id++, desc:row[ni]||'Import', amount:Math.abs(rawAmt), type:rawAmt<0?'expense':'income', cat, date, source:'Import' })
    }
    setTxs(p=>[...p,...newTxs]); setNid(id)
    setImport(false); toast(`${newTxs.length} Buchungen importiert ✓`); onBack()
  }

  return (
    <div className="fade-up">
      <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:'1.25rem' }}>
        <button onClick={onBack} style={{ color:'var(--t2)', fontSize:18 }} aria-label="Zurück">
          <i className="ti ti-arrow-left" aria-hidden="true" />
        </button>
        <span style={{ fontSize:18, fontWeight:600 }}>CSV importieren</span>
      </div>

      {headers.length === 0 ? (
        <label style={{ display:'block', border:'2px dashed var(--brd2)', borderRadius:'var(--r-xl)', padding:'2.5rem 1.5rem', textAlign:'center', cursor:'pointer' }}>
          <i className="ti ti-file-upload" style={{ fontSize:36, color:'var(--t3)', display:'block', marginBottom:10 }} aria-hidden="true" />
          <div style={{ fontSize:15, fontWeight:600, color:'var(--t1)', marginBottom:4 }}>CSV-Datei hochladen</div>
          <div style={{ fontSize:13, color:'var(--t2)' }}>Bank, Kreditkarte, PayPal, Stripe</div>
          <input type="file" accept=".csv,.txt" onChange={handleFile} style={{ display:'none' }} />
        </label>
      ) : (
        <>
          <p style={{ fontSize:13, color:'var(--t2)', marginBottom:'0.75rem' }}>Spalten zuordnen:</p>
          {[{k:'date',l:'Datum'},{k:'desc',l:'Beschreibung'},{k:'amount',l:'Betrag'}].map(({k,l}) => (
            <div key={k} style={{ display:'flex', alignItems:'center', gap:10, marginBottom:8 }}>
              <span style={{ fontSize:13, color:'var(--t2)', minWidth:110 }}>{l}</span>
              <Select value={map[k]} onChange={e=>setMap(m=>({...m,[k]:e.target.value}))} style={{ flex:1 }}>
                <option value="">– wählen –</option>
                {headers.map(h=><option key={h}>{h}</option>)}
              </Select>
            </div>
          ))}
          <div style={{ overflowX:'auto', marginTop:'1rem', marginBottom:'1rem' }}>
            <table style={{ fontSize:12, borderCollapse:'collapse', width:'100%' }}>
              <thead><tr>{headers.map(h=><th key={h} style={{ textAlign:'left', padding:'5px 8px', color:'var(--t3)', borderBottom:'1px solid var(--brd)', fontWeight:500 }}>{h}</th>)}</tr></thead>
              <tbody>{preview.map((r,i)=><tr key={i}>{r.map((v,j)=><td key={j} style={{ padding:'5px 8px', borderBottom:'1px solid var(--brd)', color:'var(--t1)' }}>{v}</td>)}</tr>)}</tbody>
            </table>
          </div>
          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={doImport} disabled={importing}>
              {importing ? '…Importiere' : `${rows.length} Zeilen importieren`}
            </Btn>
            <Btn variant="ghost" onClick={onBack}>Abbrechen</Btn>
          </div>
        </>
      )}
    </div>
  )
}

// ── Export-View ────────────────────────────────────────────────────────────

function ExportView({ txs, month, income, expense, toast, onBack }) {
  const label = month === 'all' ? 'Alle Monate' : fmtMon(month)

  const exportExcel = () => {
    const data = txs.map(t=>({ Datum:t.date.split('-').reverse().join('.'), Beschreibung:t.desc, Typ:t.type==='income'?'Einnahme':'Ausgabe', Betrag:t.type==='income'?t.amount:-t.amount, Kategorie:t.cat, Quelle:t.source }))
    const ws = XLSX.utils.json_to_sheet(data)
    ws['!cols'] = [{wch:12},{wch:30},{wch:10},{wch:12},{wch:14},{wch:12}]
    const sum  = [{ Label:'Zeitraum',Wert:label },{ Label:'Einnahmen',Wert:income },{ Label:'Ausgaben',Wert:expense },{ Label:'Saldo',Wert:income-expense }]
    const ws2  = XLSX.utils.json_to_sheet(sum)
    const wb   = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb,ws,'Buchungen')
    XLSX.utils.book_append_sheet(wb,ws2,'Zusammenfassung')
    XLSX.writeFile(wb,`Finanzen_${month||'Alle'}.xlsx`)
    toast('Excel exportiert ✓')
  }

  const exportCSV = () => {
    const hdr  = ['Datum','Beschreibung','Typ','Betrag','Kategorie','Quelle'].join(';')
    const rows = txs.map(t=>[t.date.split('-').reverse().join('.'),`"${t.desc}"`,t.type==='income'?'Einnahme':'Ausgabe',String(t.type==='income'?t.amount:-t.amount).replace('.',','),t.cat,t.source].join(';'))
    const blob = new Blob(['\uFEFF'+hdr+'\n'+rows.join('\n')],{type:'text/csv;charset=utf-8;'})
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=`Finanzen_${month||'Alle'}.csv`; a.click()
    toast('CSV exportiert ✓')
  }

  const exportPDF = () => {
    const rows = [...txs].sort((a,b)=>b.date.localeCompare(a.date)).map(t=>`<tr><td>${t.date.split('-').reverse().join('.')}</td><td>${t.desc}</td><td style="color:${t.type==='income'?'#3EC994':'#E87060'}">${t.type==='income'?'Einnahme':'Ausgabe'}</td><td style="text-align:right;font-weight:600;color:${t.type==='income'?'#3EC994':'#E87060'}">${t.type==='income'?'+':'-'}€${fmt(t.amount)}</td><td>${t.cat}</td></tr>`).join('')
    const html=`<!DOCTYPE html><html lang="de"><head><meta charset="UTF-8"><title>Finanzbericht</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:system-ui,sans-serif;padding:2rem;font-size:13px;color:#1C1B22;background:#F6F4F0}h1{font-size:24px;font-weight:700;margin-bottom:4px;color:#1C1B22}.sub{color:#6B6A78;margin-bottom:1.5rem;font-size:13px}.metrics{display:grid;grid-template-columns:1fr 1fr 1fr;gap:1rem;margin-bottom:1.5rem}.m{padding:1rem 1.1rem;border-radius:14px}.m-l{font-size:11px;font-weight:500;margin-bottom:4px;opacity:0.8}.m-v{font-size:20px;font-weight:700}table{width:100%;border-collapse:collapse}th{text-align:left;padding:8px;background:#EFEDE8;font-size:12px;font-weight:600;border-bottom:2px solid #E8E6E0}td{padding:8px;border-bottom:1px solid #E8E6E0;font-size:12px}@media print{body{padding:1rem}}</style></head><body><h1>Finanzbericht</h1><div class="sub">${label} · ${txs.length} Buchungen</div><div class="metrics"><div class="m" style="background:#E6F9F1"><div class="m-l" style="color:#3EC994">Einnahmen</div><div class="m-v" style="color:#3EC994">+€${fmt(income)}</div></div><div class="m" style="background:#FDECE8"><div class="m-l" style="color:#E87060">Ausgaben</div><div class="m-v" style="color:#E87060">-€${fmt(expense)}</div></div><div class="m" style="background:#EDEAFB"><div class="m-l" style="color:#9080D8">Saldo</div><div class="m-v" style="color:#9080D8">${income-expense>=0?'+':''}€${fmt(income-expense)}</div></div></div><table><thead><tr><th>Datum</th><th>Beschreibung</th><th>Typ</th><th style="text-align:right">Betrag</th><th>Kategorie</th></tr></thead><tbody>${rows}</tbody></table><script>setTimeout(()=>window.print(),400);<\/script></body></html>`
    const w = window.open('','_blank')
    if(w){w.document.write(html);w.document.close()}
    toast('PDF-Vorschau geöffnet')
  }

  const options = [
    { icon:'ti-file-spreadsheet', label:'Excel exportieren', sub:'.xlsx · 2 Tabellenblätter', c:'var(--mint)', bg:'var(--mint-bg)', fn:exportExcel },
    { icon:'ti-file-text', label:'CSV exportieren', sub:'Universelles Format', c:'var(--sky)', bg:'var(--sky-bg)', fn:exportCSV },
    { icon:'ti-printer', label:'PDF drucken', sub:'Gedruckter Bericht mit Zusammenfassung', c:'var(--coral)', bg:'var(--coral-bg)', fn:exportPDF },
  ]

  return (
    <div className="fade-up">
      <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:'1.25rem' }}>
        <button onClick={onBack} style={{ color:'var(--t2)', fontSize:18 }} aria-label="Zurück">
          <i className="ti ti-arrow-left" aria-hidden="true" />
        </button>
        <span style={{ fontSize:18, fontWeight:600 }}>Exportieren</span>
      </div>
      <p style={{ fontSize:13, color:'var(--t2)', marginBottom:'1.25rem' }}>
        <strong style={{ color:'var(--t1)' }}>{txs.length} Buchungen</strong> aus „{label}" exportieren
      </p>
      {options.map(o => (
        <Card key={o.label} onClick={o.fn} style={{ display:'flex', alignItems:'center', gap:14, cursor:'pointer' }}>
          <div style={{ width:44, height:44, borderRadius:'var(--r-md)', background:o.bg, color:o.c, display:'flex', alignItems:'center', justifyContent:'center', fontSize:22, flexShrink:0 }}>
            <i className={`ti ${o.icon}`} aria-hidden="true" />
          </div>
          <div>
            <div style={{ fontSize:15, fontWeight:600, color:'var(--t1)', marginBottom:2 }}>{o.label}</div>
            <div style={{ fontSize:12, color:'var(--t2)' }}>{o.sub}</div>
          </div>
        </Card>
      ))}
    </div>
  )
}

// ── Leer-Zustand ───────────────────────────────────────────────────────────

function EmptyState({ icon, text }) {
  return (
    <div style={{ textAlign:'center', padding:'2.5rem 1rem', color:'var(--t3)' }}>
      <i className={`ti ${icon}`} style={{ fontSize:32, display:'block', marginBottom:10 }} aria-hidden="true" />
      <div style={{ fontSize:14 }}>{text}</div>
    </div>
  )
}
