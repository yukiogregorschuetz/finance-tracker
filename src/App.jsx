import { useState, useEffect, useCallback, useRef } from 'react'
import { db } from './supabase.js'
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from 'recharts'
import * as XLSX from 'xlsx'
import './index.css'

// ── Konstanten ─────────────────────────────────────────────────────────────

const ICON_OPTS = ['ti-home','ti-shopping-cart','ti-car','ti-confetti','ti-heart','ti-shield','ti-trending-up','ti-dots','ti-coffee','ti-plane','ti-device-laptop','ti-gift','ti-tool','ti-school','ti-music','ti-shirt','ti-phone','ti-building','ti-wallet','ti-cash','ti-star','ti-bike','ti-swim']
const COLOR_OPTS = ['sky','mint','violet','rose','coral','t2','t3']
const COLOR_MAP = {
  sky:    { c:'var(--sky)',    bg:'var(--sky-bg)',    hex:'#50A8E0', hexDk:'#70BCEC' },
  mint:   { c:'var(--mint)',   bg:'var(--mint-bg)',   hex:'#3EC994', hexDk:'#52C898' },
  violet: { c:'var(--violet)', bg:'var(--violet-bg)', hex:'#9080D8', hexDk:'#A898E0' },
  rose:   { c:'var(--rose)',   bg:'var(--rose-bg)',   hex:'#E86880', hexDk:'#F08898' },
  coral:  { c:'var(--coral)',  bg:'var(--coral-bg)',  hex:'#E87060', hexDk:'#F08878' },
  t2:     { c:'var(--t2)',     bg:'var(--card2)',     hex:'#6B6A78', hexDk:'#8A8898' },
  t3:     { c:'var(--t3)',     bg:'var(--card3)',     hex:'#A4A3B0', hexDk:'#52506A' },
}

const DEFAULT_CATS = [
  { name:'Wohnen',       icon:'ti-home',         color:'sky'    },
  { name:'Lebensmittel', icon:'ti-shopping-cart', color:'mint'   },
  { name:'Transport',    icon:'ti-car',           color:'violet' },
  { name:'Freizeit',     icon:'ti-confetti',      color:'rose'   },
  { name:'Gesundheit',   icon:'ti-heart',         color:'coral'  },
  { name:'Versicherung', icon:'ti-shield',        color:'t2'     },
  { name:'Gehalt',       icon:'ti-trending-up',   color:'mint'   },
  { name:'Sonstiges',    icon:'ti-dots',          color:'t3'     },
]
const DEFAULT_SOURCES = ['Bankkonto','Kreditkarte','PayPal','Stripe','Bar']
const DEFAULT_RULES = {
  Lebensmittel: 'rewe,lidl,aldi,edeka,penny,netto,kaufland,dm,rossmann,bäcker,metzger,wasgau,backerei,backshop,drogerie,müller,mueller,globus',
  Wohnen:       'miete,nebenkosten,strom,gas,wasser,internet,telefon,hausgeld,1und1,1+1,vodafone,telekom,stadtwerke,verbandsgemeinde',
  Transport:    'monatskarte,db,bahn,öpnv,tankstelle,shell,aral,bp,uber,taxi,flixbus,lufthansa,minera,kraftstoff,parkgebühr,parkgebuehr,hyundai,kfz',
  Freizeit:     'netflix,spotify,disney,kino,theater,konzert,steam,playstation,gym,fitnessstudio,fitness,supercell,riot games,epic games,google play,kinoheld,restaurant,pizzeria,ristorante,gastro,diner,eiscafe',
  Gesundheit:   'apotheke,arzt,krankenhaus,optiker,zahnarzt,physiotherapie,rats-apotheke',
  Versicherung: 'versicherung,allianz,aok,tkk,barmer,huk,continentale,europa verbund,gewerkschaft,verdi',
  Gehalt:       'gehalt,lohn,freelance,honorar,bundeskasse,renten service,rente,pension,zuwendung',
  Sonstiges:    'amazon,topstep,consors,easycredit,teambank,paypal,microsoft,google,anthropic,claude',
}
const DEF_BUDGETS   = { Wohnen:900, Lebensmittel:400, Transport:100, Freizeit:150, Gesundheit:100, Versicherung:150 }
const DEF_RECURRING = [
  { id:1, desc:'Miete',            amount:850,   type:'expense', cat:'Wohnen',       source:'Bankkonto', freq:'Monatlich', nextDate:'2026-06-01' },
  { id:2, desc:'Netflix',          amount:17.99, type:'expense', cat:'Freizeit',     source:'PayPal',    freq:'Monatlich', nextDate:'2026-06-05' },
  { id:3, desc:'Kfz-Versicherung', amount:112,   type:'expense', cat:'Versicherung', source:'Bankkonto', freq:'Monatlich', nextDate:'2026-06-10' },
]
const FREQS = ['Monatlich','Wöchentlich','Quartalsweise','Jährlich']

// ── Helpers ────────────────────────────────────────────────────────────────

const fmt    = n  => n.toLocaleString('de-DE',{ minimumFractionDigits:2, maximumFractionDigits:2 })
const today  = () => new Date().toISOString().split('T')[0]
const MN     = ['','Jan','Feb','Mär','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Dez']
const fmtMon = m  => MN[parseInt(m.slice(5))] + ' ' + m.slice(2,4)

function getCatInfo(catName, cats) {
  const cat = (cats || DEFAULT_CATS).find(c => c.name === catName) || { icon:'ti-dots', color:'t3' }
  return { ...(COLOR_MAP[cat.color] || COLOR_MAP.t3), icon: cat.icon }
}

function aiCat(desc, rules) {
  const low = desc.toLowerCase()
  for (const [cat, keywords] of Object.entries(rules || {})) {
    const kws = typeof keywords === 'string' ? keywords.split(',').map(k => k.trim()) : (keywords || [])
    if (kws.some(k => k && low.includes(k))) return cat
  }
  return 'Sonstiges'
}

// ── Supabase ───────────────────────────────────────────────────────────────

async function syncToSupabase(txs, budgets, recurring) {
  try {
    await db.from('transactions').delete().neq('id', 0)
    if (txs.length > 0) {
      await db.from('transactions').insert(
        txs.map(t => ({ description:t.desc, amount:t.amount, type:t.type, cat:t.cat, date:t.date, source:t.source }))
      )
    }
    const budgetRows = Object.entries(budgets).map(([cat, amount]) => ({ cat, amount }))
    if (budgetRows.length > 0) await db.from('budgets').upsert(budgetRows, { onConflict:'cat' })
  } catch(e) { console.log('Sync fehler:', e) }
}

async function loadFromSupabase() {
  try {
    const [txRes, budRes] = await Promise.all([
      db.from('transactions').select('*').order('date', { ascending:false }),
      db.from('budgets').select('*')
    ])
    const txs = (txRes.data || []).map(t => ({ id:t.id, desc:t.description, amount:t.amount, type:t.type, cat:t.cat, date:t.date, source:t.source }))
    const budgets = {}
    ;(budRes.data || []).forEach(b => { budgets[b.cat] = b.amount })
    return { txs, budgets }
  } catch(e) { return null }
}

// ── Hooks ──────────────────────────────────────────────────────────────────

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

// ── Shared UI ──────────────────────────────────────────────────────────────

function Card({ children, style={}, onClick }) {
  return (
    <div onClick={onClick} style={{
      background:'var(--card)', borderRadius:'var(--r-lg)',
      border:'1px solid var(--brd)', boxShadow:'var(--sh-sm)',
      padding:'1.1rem 1.25rem', marginBottom:'0.75rem',
      cursor: onClick ? 'pointer' : 'default', ...style
    }}>{children}</div>
  )
}

function Pill({ children, color, bg }) {
  return (
    <span style={{
      display:'inline-flex', alignItems:'center', gap:4,
      background:bg, color:color, fontSize:11, fontWeight:500,
      padding:'3px 9px', borderRadius:100, whiteSpace:'nowrap'
    }}>{children}</span>
  )
}

function CatIcon({ cat, cats, size=34 }) {
  const info = getCatInfo(cat, cats)
  return (
    <div style={{
      width:size, height:size, borderRadius:'50%',
      background:info.bg, color:info.c, flexShrink:0,
      display:'flex', alignItems:'center', justifyContent:'center', fontSize:size*0.44
    }}>
      <i className={`ti ${info.icon}`} aria-hidden="true" />
    </div>
  )
}

function Btn({ children, variant='primary', onClick, disabled, style={} }) {
  const base = {
    display:'inline-flex', alignItems:'center', gap:6,
    padding:'9px 18px', borderRadius:'var(--r-md)', fontSize:14,
    fontWeight:500, cursor: disabled ? 'not-allowed' : 'pointer',
    opacity: disabled ? 0.6 : 1, border:'1px solid transparent', ...style
  }
  const v = {
    primary: { background:'var(--t1)', color:'var(--page)' },
    ghost:   { background:'var(--card2)', color:'var(--t1)', border:'1px solid var(--brd)' },
    danger:  { background:'var(--coral-bg)', color:'var(--coral)' },
  }
  return <button style={{ ...base, ...v[variant] }} onClick={onClick} disabled={disabled}>{children}</button>
}

function Field({ label, children, style={} }) {
  return (
    <div style={{ display:'flex', flexDirection:'column', gap:5, ...style }}>
      <label style={{ fontSize:12, fontWeight:500, color:'var(--t2)' }}>{label}</label>
      {children}
    </div>
  )
}

const inp = {
  width:'100%', height:40, padding:'0 12px',
  borderRadius:'var(--r-md)', border:'1px solid var(--brd2)',
  background:'var(--card)', color:'var(--t1)', fontSize:14,
}

function Input({ style, ...props }) {
  return <input style={{ ...inp, ...style }} {...props} />
}

function Sel({ children, style, ...props }) {
  return (
    <select style={{ ...inp, ...style, cursor:'pointer' }} {...props}>
      {children}
    </select>
  )
}

function EmptyState({ icon, text }) {
  return (
    <div style={{ textAlign:'center', padding:'2.5rem 1rem', color:'var(--t3)' }}>
      <i className={`ti ${icon}`} style={{ fontSize:32, display:'block', marginBottom:10 }} aria-hidden="true" />
      <div style={{ fontSize:14 }}>{text}</div>
    </div>
  )
}

// ── App ────────────────────────────────────────────────────────────────────

export default function App() {
  const [tab, setTab]           = useState('dashboard')
  const [txs, setTxs]           = useLS('ft-txs', [])
  const [budgets, setBudgets]   = useLS('ft-budgets', DEF_BUDGETS)
  const [recurring, setRec]     = useLS('ft-recurring', DEF_RECURRING)
  const [month, setMonth]       = useLS('ft-month', new Date().toISOString().slice(0,7))
  const [nid, setNid]           = useLS('ft-nid', 200)
  const [cats, setCats]         = useLS('ft-cats', DEFAULT_CATS)
  const [sources, setSources]   = useLS('ft-sources', DEFAULT_SOURCES)
  const [rules, setRules]       = useLS('ft-rules', DEFAULT_RULES)
  const loaded                  = useRef(false)

  useEffect(() => {
    loadFromSupabase().then(data => {
      if (data?.txs?.length > 0) setTxs(data.txs)
      if (Object.keys(data?.budgets || {}).length > 0) setBudgets(data.budgets)
      loaded.current = true
    })
  }, [])

  useEffect(() => {
    if (!loaded.current) return
    syncToSupabase(txs, budgets, recurring)
  }, [txs, budgets, recurring])

  const [theme, toggleTheme] = useTheme()
  const [toast, setToast]    = useState(null)
  const showToast = useCallback((msg, type='ok') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 2600)
  }, [])

  const months   = [...new Set(txs.map(t => t.date.slice(0,7)))].sort().reverse()
  const filtered = txs.filter(t => month === 'all' || t.date.startsWith(month))
  const income   = filtered.filter(t => t.type==='income').reduce((s,t)=>s+t.amount, 0)
  const expense  = filtered.filter(t => t.type==='expense').reduce((s,t)=>s+t.amount, 0)
  const catSpend = {}
  filtered.filter(t => t.type==='expense').forEach(t => { catSpend[t.cat]=(catSpend[t.cat]||0)+t.amount })
  const warnings = Object.entries(budgets).filter(([c,lim]) => lim>0 && (catSpend[c]||0)>=lim*0.8)

  const NAV = [
    { id:'dashboard',    icon:'ti-layout-dashboard', label:'Übersicht'    },
    { id:'transactions', icon:'ti-list-details',     label:'Buchungen'    },
    { id:'budget',       icon:'ti-target',           label:'Budget'       },
    { id:'recurring',    icon:'ti-repeat',           label:'Fixkosten'    },
    { id:'settings',     icon:'ti-settings',         label:'Einstellungen'},
  ]

  return (
    <div style={{ display:'flex', flexDirection:'column', height:'100%', background:'var(--page)' }}>
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
        {tab !== 'settings' && (
          <Sel value={month} onChange={e=>setMonth(e.target.value)}
            style={{ width:'auto', height:34, fontSize:13, paddingRight:28 }}>
            <option value="all">Alle Monate</option>
            {months.map(m=><option key={m} value={m}>{fmtMon(m)}</option>)}
          </Sel>
        )}
        <button onClick={toggleTheme} style={{
          width:34, height:34, borderRadius:'var(--r-md)',
          background:'var(--card2)', border:'1px solid var(--brd)',
          display:'flex', alignItems:'center', justifyContent:'center',
          color:'var(--t2)', fontSize:16, flexShrink:0
        }} aria-label="Design wechseln">
          <i className={`ti ${theme==='dark'?'ti-sun':'ti-moon'}`} aria-hidden="true" />
        </button>
      </header>

      {warnings.length > 0 && tab !== 'settings' && (
        <div style={{ background:'var(--honey-bg)', borderBottom:'1px solid var(--honey-dim)', padding:'10px 1rem' }}>
          {warnings.map(([cat,lim]) => {
            const spent=catSpend[cat]||0, pct=Math.round(spent/lim*100), over=spent>lim
            return (
              <div key={cat} style={{ fontSize:13, color:'var(--honey)', display:'flex', alignItems:'center', gap:6, marginBottom:2 }}>
                <i className="ti ti-alert-triangle" aria-hidden="true" />
                {over ? `Budget überschritten: ${cat} · €${fmt(spent)} / €${fmt(lim)}` : `Budget fast voll: ${cat} · ${pct}%`}
              </div>
            )
          })}
        </div>
      )}

      <main style={{ flex:1, overflowY:'auto', paddingBottom:`calc(var(--nav-h) + 8px)` }}>
        <div style={{ padding:'1rem', maxWidth:680, margin:'0 auto' }}>
          {tab==='dashboard'    && <Dashboard filtered={filtered} income={income} expense={expense} catSpend={catSpend} allTxs={txs} theme={theme} cats={cats} />}
          {tab==='transactions' && <Transactions txs={filtered} setTxs={setTxs} nid={nid} setNid={setNid} toast={showToast} cats={cats} sources={sources} rules={rules} />}
          {tab==='budget'       && <Budget budgets={budgets} setBudgets={setBudgets} catSpend={catSpend} toast={showToast} cats={cats} />}
          {tab==='recurring'    && <Recurring recurring={recurring} setRec={setRec} setTxs={setTxs} nid={nid} setNid={setNid} toast={showToast} cats={cats} sources={sources} />}
          {tab==='settings' && <Settings cats={cats} setCats={setCats} sources={sources} setSources={setSources} rules={rules} setRules={setRules} budgets={budgets} setBudgets={setBudgets} setTxs={setTxs} setRec={setRec} toast={showToast} txs={txs} income={income} expense={expense} month={month} nid={nid} setNid={setNid} />}
        </div>
      </main>

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
            <button key={n.id} onClick={()=>setTab(n.id)} style={{
              flex:1, display:'flex', flexDirection:'column', alignItems:'center',
              gap:3, padding:'6px 0', border:'none', background:'none', cursor:'pointer',
              color: active ? 'var(--violet)' : 'var(--t3)', transition:'color 0.2s',
            }}>
              <div style={{
                width:38, height:28, borderRadius:14, display:'flex', alignItems:'center',
                justifyContent:'center', fontSize:18,
                background: active ? 'var(--violet-bg)' : 'transparent', transition:'background 0.2s',
              }}>
                <i className={`ti ${n.icon}`} aria-hidden="true" />
              </div>
              <span style={{ fontSize:10, fontWeight: active?600:400 }}>{n.label}</span>
            </button>
          )
        })}
      </nav>

      {toast && (
        <div style={{
          position:'fixed', bottom:`calc(var(--nav-h) + 12px)`, left:'50%',
          transform:'translateX(-50%)', zIndex:200,
          background: toast.type==='err' ? 'var(--coral)' : 'var(--t1)',
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

function Dashboard({ filtered, income, expense, catSpend, allTxs, theme, cats }) {
  const balance = income - expense
  const dk      = theme === 'dark'

  const catData = Object.entries(catSpend)
    .map(([name,value]) => ({ name, value: Math.round(value*100)/100 }))
    .sort((a,b) => b.value - a.value)

  const monthMap = {}
  allTxs.forEach(t => {
    const m = t.date.slice(0,7)
    if (!monthMap[m]) monthMap[m] = { name:fmtMon(m), inc:0, exp:0 }
    if (t.type==='income') monthMap[m].inc += t.amount
    else monthMap[m].exp += t.amount
  })
  const monthData = Object.values(monthMap)
    .sort((a,b) => a.name.localeCompare(b.name))
    .map(d => ({ ...d, inc:Math.round(d.inc), exp:Math.round(d.exp) }))

  const recentTxs = [...filtered].sort((a,b)=>b.date.localeCompare(a.date)).slice(0,5)

  const metrics = [
    { label:'Einnahmen', val:`+€${fmt(income)}`,                             c:'var(--mint)',   bg:'var(--mint-bg)'   },
    { label:'Ausgaben',  val:`-€${fmt(expense)}`,                            c:'var(--coral)',  bg:'var(--coral-bg)'  },
    { label:'Saldo',     val:`${balance>=0?'+':''}€${fmt(balance)}`,         c:'var(--violet)', bg:'var(--violet-bg)' },
    { label:'Buchungen', val:filtered.length,                                c:'var(--sky)',    bg:'var(--sky-bg)'    },
  ]

  return (
    <div className="fade-up">
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'0.6rem', marginBottom:'1.25rem' }}>
        {metrics.map(m => (
          <div key={m.label} style={{ background:m.bg, borderRadius:'var(--r-lg)', padding:'1rem 1.1rem' }}>
            <div style={{ fontSize:12, fontWeight:500, color:m.c, marginBottom:6, opacity:0.8 }}>{m.label}</div>
            <div style={{ fontSize:22, fontWeight:600, color:m.c, letterSpacing:'-0.5px', fontVariantNumeric:'tabular-nums' }}>{m.val}</div>
          </div>
        ))}
      </div>

      <Card>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:'1rem' }}>Ausgaben nach Kategorie</div>
        {catData.length===0
          ? <EmptyState icon="ti-chart-bar" text="Keine Ausgaben im Zeitraum" />
          : <ResponsiveContainer width="100%" height={180}>
              <BarChart data={catData} margin={{ top:0, right:0, left:-20, bottom:0 }}>
                <XAxis dataKey="name" tick={{ fontSize:10, fill:'var(--t3)' }} />
                <YAxis tick={{ fontSize:10, fill:'var(--t3)' }} tickFormatter={v=>'€'+v} />
                <Tooltip formatter={v=>[`€${fmt(v)}`,'Ausgaben']} contentStyle={{ fontSize:12, borderRadius:10, border:'none', background:'var(--card)', color:'var(--t1)' }} cursor={{ fill:'var(--card2)' }} />
                <Bar dataKey="value" radius={[6,6,0,0]}>
                  {catData.map(e => { const info=getCatInfo(e.name,cats); return <Cell key={e.name} fill={dk?info.hexDk:info.hex} /> })}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
        }
      </Card>

      <Card>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:'0.6rem' }}>Monatsverlauf</div>
        <div style={{ display:'flex', gap:16, marginBottom:'0.75rem' }}>
          {[{col:dk?'#52C898':'#3EC994',label:'Einnahmen'},{col:dk?'#F08878':'#E87060',label:'Ausgaben'}].map(l=>(
            <span key={l.label} style={{ display:'flex', alignItems:'center', gap:5, fontSize:12, color:'var(--t2)' }}>
              <span style={{ width:10, height:10, borderRadius:3, background:l.col }} />{l.label}
            </span>
          ))}
        </div>
        <ResponsiveContainer width="100%" height={160}>
          <BarChart data={monthData} margin={{ top:0, right:0, left:-20, bottom:0 }}>
            <XAxis dataKey="name" tick={{ fontSize:10, fill:'var(--t3)' }} />
            <YAxis tick={{ fontSize:10, fill:'var(--t3)' }} tickFormatter={v=>'€'+v} />
            <Tooltip formatter={(v,n)=>[`€${fmt(v)}`,n==='inc'?'Einnahmen':'Ausgaben']} contentStyle={{ fontSize:12, borderRadius:10, border:'none', background:'var(--card)', color:'var(--t1)' }} cursor={{ fill:'var(--card2)' }} />
            <Bar dataKey="inc" fill={dk?'#52C898':'#3EC994'} radius={[4,4,0,0]} />
            <Bar dataKey="exp" fill={dk?'#F08878':'#E87060'} radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>

      <Card>
        <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:'0.75rem' }}>Letzte Buchungen</div>
        {recentTxs.length===0
          ? <EmptyState icon="ti-receipt" text="Keine Buchungen" />
          : recentTxs.map((t,i) => (
              <div key={t.id} style={{ display:'flex', alignItems:'center', gap:10, padding:'8px 0', borderTop:i>0?'1px solid var(--brd)':'none' }}>
                <CatIcon cat={t.cat} cats={cats} size={32} />
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{t.desc}</div>
                  <div style={{ fontSize:11, color:'var(--t3)' }}>{t.date.split('-').reverse().join('.')}</div>
                </div>
                <div style={{ fontSize:14, fontWeight:600, color:t.type==='income'?'var(--mint)':'var(--coral)', fontVariantNumeric:'tabular-nums' }}>
                  {t.type==='income'?'+':'-'}€{fmt(t.amount)}
                </div>
              </div>
            ))
        }
      </Card>
    </div>
  )
}

// ── Transactions ───────────────────────────────────────────────────────────

function Transactions({ txs, setTxs, nid, setNid, toast, cats, sources, rules }) {
  const catNames = cats.map(c=>c.name)
  const BLANK = { desc:'', amount:'', type:'expense', cat:catNames[0]||'Sonstiges', date:today(), source:sources[0]||'Bankkonto' }
  const [form, setForm]     = useState(null)
  const [aiLoad, setAiLoad] = useState(false)
  const [typeF, setTypeF]   = useState('all')
  const [catF, setCatF]     = useState('all')
  const [q, setQ]           = useState('')
  const set = k => e => setForm(p=>({ ...p, [k]:e.target.value }))

  const sorted = [...txs]
    .filter(t => (typeF==='all'||t.type===typeF) && (catF==='all'||t.cat===catF) && (!q||t.desc.toLowerCase().includes(q.toLowerCase())))
    .sort((a,b)=>b.date.localeCompare(a.date))

  const handleAI = () => {
    if (!form?.desc) { toast('Beschreibung eingeben','err'); return }
    setAiLoad(true)
    const cat = aiCat(form.desc, rules)
    setForm(p=>({ ...p, cat })); toast('Kategorie: '+cat)
    setAiLoad(false)
  }

  const save = () => {
    if (!form.desc||!form.amount||!form.date) { toast('Pflichtfelder ausfüllen','err'); return }
    setTxs(p=>[{ ...form, id:nid, amount:parseFloat(form.amount) }, ...p])
    setNid(n=>n+1); setForm(null); toast('Buchung gespeichert ✓')
  }

  const del = id => { setTxs(p=>p.filter(t=>t.id!==id)); toast('Gelöscht') }

  return (
    <div className="fade-up">
      <div style={{ display:'flex', gap:8, marginBottom:'0.75rem', flexWrap:'wrap' }}>
        <Sel value={typeF} onChange={e=>setTypeF(e.target.value)} style={{ flex:1, minWidth:100, height:38 }}>
          <option value="all">Alle Typen</option>
          <option value="income">Einnahmen</option>
          <option value="expense">Ausgaben</option>
        </Sel>
        <Sel value={catF} onChange={e=>setCatF(e.target.value)} style={{ flex:1, minWidth:120, height:38 }}>
          <option value="all">Alle Kategorien</option>
          {catNames.map(c=><option key={c}>{c}</option>)}
        </Sel>
      </div>
      <div style={{ display:'flex', gap:8, marginBottom:'1rem' }}>
        <Input value={q} onChange={e=>setQ(e.target.value)} placeholder="Suchen…" style={{ flex:1, height:38 }} />
        <Btn onClick={()=>setForm(BLANK)}><i className="ti ti-plus" aria-hidden="true" /> Neu</Btn>
      </div>

      {form && (
        <div style={{ background:'var(--card2)', borderRadius:'var(--r-lg)', padding:'1.1rem', marginBottom:'1rem', border:'1px solid var(--brd)' }}>
          <Field label="Beschreibung" style={{ marginBottom:10 }}>
            <div style={{ display:'flex', gap:6 }}>
              <Input value={form.desc} onChange={set('desc')} placeholder="z.B. REWE, Miete…" style={{ flex:1 }} />
              <button onClick={handleAI} disabled={aiLoad} style={{
                height:40, padding:'0 12px', borderRadius:'var(--r-md)',
                background:'var(--violet-bg)', color:'var(--violet)',
                border:'1px solid var(--violet-dim)', fontSize:13, fontWeight:500, cursor:'pointer', flexShrink:0
              }}>{aiLoad?'…':'🤖 KI'}</button>
            </div>
          </Field>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:10 }}>
            <Field label="Betrag (€)">
              <Input type="number" value={form.amount} onChange={set('amount')} placeholder="0.00" step="0.01" />
            </Field>
            <Field label="Typ">
              <Sel value={form.type} onChange={set('type')}>
                <option value="expense">Ausgabe</option>
                <option value="income">Einnahme</option>
              </Sel>
            </Field>
            <Field label="Kategorie">
              <Sel value={form.cat} onChange={set('cat')}>
                {catNames.map(c=><option key={c}>{c}</option>)}
              </Sel>
            </Field>
            <Field label="Datum">
              <Input type="date" value={form.date} onChange={set('date')} />
            </Field>
            <Field label="Quelle">
              <Sel value={form.source} onChange={set('source')}>
                {sources.map(s=><option key={s}>{s}</option>)}
              </Sel>
            </Field>
          </div>
          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={save}>Speichern</Btn>
            <Btn variant="ghost" onClick={()=>setForm(null)}>Abbrechen</Btn>
          </div>
        </div>
      )}

      {sorted.length===0
        ? <EmptyState icon="ti-receipt" text="Keine Buchungen gefunden" />
        : sorted.map((t,i) => {
            const info = getCatInfo(t.cat, cats)
            return (
              <div key={t.id} style={{ display:'flex', alignItems:'center', gap:10, padding:'10px 0', borderTop:i>0?'1px solid var(--brd)':'none' }}>
                <CatIcon cat={t.cat} cats={cats} size={36} />
                <div style={{ flex:1, minWidth:0 }}>
                  <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{t.desc}</div>
                  <div style={{ display:'flex', gap:6, marginTop:3, flexWrap:'wrap' }}>
                    <Pill color={info.c} bg={info.bg}>{t.cat}</Pill>
                    <span style={{ fontSize:11, color:'var(--t3)' }}>{t.date.split('-').reverse().join('.')}</span>
                    <span style={{ fontSize:11, color:'var(--t3)' }}>{t.source}</span>
                  </div>
                </div>
                <div style={{ display:'flex', flexDirection:'column', alignItems:'flex-end', gap:4 }}>
                  <div style={{ fontSize:15, fontWeight:600, color:t.type==='income'?'var(--mint)':'var(--coral)', fontVariantNumeric:'tabular-nums' }}>
                    {t.type==='income'?'+':'-'}€{fmt(t.amount)}
                  </div>
                  <button onClick={()=>del(t.id)} style={{ fontSize:13, color:'var(--t3)', background:'none', border:'none', cursor:'pointer', padding:0 }} aria-label="Löschen">
                    <i className="ti ti-trash" />
                  </button>
                </div>
              </div>
            )
          })
      }
    </div>
  )
}

// ── Budget ─────────────────────────────────────────────────────────────────

function Budget({ budgets, setBudgets, catSpend, toast, cats }) {
  const expCats = cats.filter(c=>c.name!=='Gehalt').map(c=>c.name)
  const totalBudget = expCats.reduce((s,c)=>s+(budgets[c]||0), 0)
  const totalSpend  = expCats.reduce((s,c)=>s+(catSpend[c]||0), 0)

  return (
    <div className="fade-up">
      <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:'0.6rem', marginBottom:'1.25rem' }}>
        {[
          { label:'Geplant',     val:`€${fmt(totalBudget)}`, c:'var(--violet)', bg:'var(--violet-bg)' },
          { label:'Ausgegeben',  val:`€${fmt(totalSpend)}`,  c:'var(--coral)',  bg:'var(--coral-bg)'  },
        ].map(m => (
          <div key={m.label} style={{ background:m.bg, borderRadius:'var(--r-lg)', padding:'1rem 1.1rem' }}>
            <div style={{ fontSize:12, fontWeight:500, color:m.c, marginBottom:6, opacity:0.8 }}>{m.label}</div>
            <div style={{ fontSize:22, fontWeight:600, color:m.c, fontVariantNumeric:'tabular-nums' }}>{m.val}</div>
          </div>
        ))}
      </div>

      {expCats.map(cat => {
        const lim   = budgets[cat]||0
        const spent = catSpend[cat]||0
        const pct   = lim>0 ? Math.min(spent/lim*100,100) : 0
        const over  = spent>lim && lim>0
        const info  = getCatInfo(cat, cats)
        return (
          <Card key={cat}>
            <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:lim>0?10:0 }}>
              <CatIcon cat={cat} cats={cats} size={32} />
              <div style={{ flex:1 }}>
                <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)' }}>{cat}</div>
                <div style={{ fontSize:12, color:'var(--t3)' }}>€{fmt(spent)} / €{fmt(lim)}</div>
              </div>
              <div style={{ display:'flex', alignItems:'center', gap:6 }}>
                <span style={{ fontSize:13, color:over?'var(--coral)':'var(--t2)', fontWeight:500 }}>
                  {lim>0?`${Math.round(spent/lim*100)}%`:'–'}
                </span>
                <input type="number" value={lim||''} onChange={e=>setBudgets(p=>({ ...p, [cat]:parseFloat(e.target.value)||0 }))}
                  placeholder="Budget" style={{ ...inp, width:90, height:34, fontSize:13, textAlign:'right' }} />
              </div>
            </div>
            {lim>0 && (
              <div style={{ height:6, background:'var(--card2)', borderRadius:99, overflow:'hidden' }}>
                <div style={{ height:'100%', width:`${pct}%`, background:over?'var(--coral)':info.c, borderRadius:99, transition:'width 0.4s' }} />
              </div>
            )}
          </Card>
        )
      })}
    </div>
  )
}

// ── Recurring ──────────────────────────────────────────────────────────────

function Recurring({ recurring, setRec, setTxs, nid, setNid, toast, cats, sources }) {
  const catNames = cats.map(c=>c.name)
  const BLANK = { desc:'', amount:'', type:'expense', cat:catNames[0]||'Sonstiges', source:sources[0]||'Bankkonto', freq:'Monatlich', nextDate:today() }
  const [form, setForm] = useState(null)
  const set = k => e => setForm(p=>({ ...p, [k]:e.target.value }))

  const addAll = () => {
    const t0=today(); let id=nid
    const newTxs = recurring.map(r=>({ id:id++, desc:r.desc, amount:r.amount, type:r.type, cat:r.cat, date:t0, source:r.source }))
    setTxs(p=>[...newTxs,...p]); setNid(id)
    toast(`${newTxs.length} Buchungen hinzugefügt ✓`)
  }

  const save = () => {
    if (!form.desc||!form.amount) { toast('Pflichtfelder ausfüllen','err'); return }
    setRec(p=>[...p, { ...form, id:nid, amount:parseFloat(form.amount) }])
    setNid(n=>n+1); setForm(null); toast('Fixkosten gespeichert ✓')
  }

  return (
    <div className="fade-up">
      <div style={{ display:'flex', gap:8, marginBottom:'1rem' }}>
        <Btn onClick={addAll} style={{ flex:1 }}><i className="ti ti-plus" /> Alle jetzt buchen</Btn>
        <Btn variant="ghost" onClick={()=>setForm(BLANK)}><i className="ti ti-plus" /> Neu</Btn>
      </div>

      {form && (
        <div style={{ background:'var(--card2)', borderRadius:'var(--r-lg)', padding:'1.1rem', marginBottom:'1rem', border:'1px solid var(--brd)' }}>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:10 }}>
            <Field label="Beschreibung">
              <Input value={form.desc} onChange={set('desc')} placeholder="z.B. Miete" />
            </Field>
            <Field label="Betrag (€)">
              <Input type="number" value={form.amount} onChange={set('amount')} placeholder="0.00" step="0.01" />
            </Field>
            <Field label="Typ">
              <Sel value={form.type} onChange={set('type')}>
                <option value="expense">Ausgabe</option>
                <option value="income">Einnahme</option>
              </Sel>
            </Field>
            <Field label="Kategorie">
              <Sel value={form.cat} onChange={set('cat')}>
                {catNames.map(c=><option key={c}>{c}</option>)}
              </Sel>
            </Field>
            <Field label="Quelle">
              <Sel value={form.source} onChange={set('source')}>
                {sources.map(s=><option key={s}>{s}</option>)}
              </Sel>
            </Field>
            <Field label="Häufigkeit">
              <Sel value={form.freq} onChange={set('freq')}>
                {FREQS.map(f=><option key={f}>{f}</option>)}
              </Sel>
            </Field>
          </div>
          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={save}>Speichern</Btn>
            <Btn variant="ghost" onClick={()=>setForm(null)}>Abbrechen</Btn>
          </div>
        </div>
      )}

      {recurring.length===0
        ? <EmptyState icon="ti-repeat" text="Keine Fixkosten" />
        : recurring.map(r => {
            const info = getCatInfo(r.cat, cats)
            return (
              <Card key={r.id}>
                <div style={{ display:'flex', alignItems:'center', gap:10 }}>
                  <CatIcon cat={r.cat} cats={cats} size={36} />
                  <div style={{ flex:1, minWidth:0 }}>
                    <div style={{ fontSize:14, fontWeight:500, color:'var(--t1)' }}>{r.desc}</div>
                    <div style={{ display:'flex', gap:6, marginTop:3 }}>
                      <Pill color={info.c} bg={info.bg}>{r.cat}</Pill>
                      <span style={{ fontSize:11, color:'var(--t3)' }}>{r.freq}</span>
                    </div>
                  </div>
                  <div style={{ display:'flex', flexDirection:'column', alignItems:'flex-end', gap:4 }}>
                    <div style={{ fontSize:15, fontWeight:600, color:r.type==='income'?'var(--mint)':'var(--coral)', fontVariantNumeric:'tabular-nums' }}>
                      {r.type==='income'?'+':'-'}€{fmt(r.amount)}
                    </div>
                    <button onClick={()=>setRec(p=>p.filter(x=>x.id!==r.id))} style={{ fontSize:13, color:'var(--t3)', background:'none', border:'none', cursor:'pointer', padding:0 }} aria-label="Löschen">
                      <i className="ti ti-trash" />
                    </button>
                  </div>
                </div>
              </Card>
            )
          })
      }
    </div>
  )
}

// ── Settings ───────────────────────────────────────────────────────────────

function Settings({ cats, setCats, sources, setSources, rules, setRules, budgets, setBudgets, setTxs, setRec, toast, txs, income, expense, month, nid, setNid }) {
  const [section, setSection]   = useState('cats')
  const [newCatName, setNewCatName] = useState('')
  const [newCatIcon, setNewCatIcon] = useState('ti-dots')
  const [newCatColor, setNewCatColor] = useState('mint')
  const [newSource, setNewSource]   = useState('')
  const [editRules, setEditRules]   = useState(() =>
    Object.fromEntries(cats.map(c => [c.name, typeof rules[c.name]==='string' ? rules[c.name] : (rules[c.name]||[]).join(',')]))
  )

  const addCat = () => {
    if (!newCatName.trim()) { toast('Name eingeben','err'); return }
    if (cats.find(c=>c.name===newCatName.trim())) { toast('Existiert bereits','err'); return }
    const name = newCatName.trim()
    setCats(p=>[...p, { name, icon:newCatIcon, color:newCatColor }])
    setRules(p=>({ ...p, [name]:'' }))
    setEditRules(p=>({ ...p, [name]:'' }))
    setNewCatName(''); toast('Kategorie erstellt ✓')
  }

  const delCat = name => {
    if (name==='Sonstiges') { toast('Sonstiges kann nicht gelöscht werden','err'); return }
    setCats(p=>p.filter(c=>c.name!==name))
    const r2={...rules}; delete r2[name]; setRules(r2)
    toast('Kategorie gelöscht')
  }

  const saveRules = () => { setRules(editRules); toast('KI-Regeln gespeichert ✓') }

  const addSource = () => {
    if (!newSource.trim()) { toast('Name eingeben','err'); return }
    setSources(p=>[...p, newSource.trim()])
    setNewSource(''); toast('Quelle hinzugefügt ✓')
  }

  const exportExcel = () => {
    const label = month==='all'?'Alle Monate':fmtMon(month)
    const data = txs.map(t=>({ Datum:t.date.split('-').reverse().join('.'), Beschreibung:t.desc, Typ:t.type==='income'?'Einnahme':'Ausgabe', Betrag:t.type==='income'?t.amount:-t.amount, Kategorie:t.cat, Quelle:t.source }))
    const ws = XLSX.utils.json_to_sheet(data)
    ws['!cols'] = [{wch:12},{wch:30},{wch:10},{wch:12},{wch:14},{wch:12}]
    const sum = [{ Label:'Zeitraum',Wert:label },{ Label:'Einnahmen',Wert:income },{ Label:'Ausgaben',Wert:expense },{ Label:'Saldo',Wert:income-expense }]
    const ws2 = XLSX.utils.json_to_sheet(sum)
    const wb  = XLSX.utils.book_new()
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

  const resetAll = () => {
    if (!window.confirm('Alle Daten zurücksetzen? Kann nicht rückgängig gemacht werden.')) return
    setTxs([]); setRec(DEF_RECURRING); setBudgets(DEF_BUDGETS)
    setCats(DEFAULT_CATS); setSources(DEFAULT_SOURCES); setRules(DEFAULT_RULES)
    toast('Zurückgesetzt')
  }

  const SECTIONS = [
    { id:'cats',    label:'Kategorien', icon:'ti-tag'        },
    { id:'sources', label:'Quellen',    icon:'ti-credit-card'},
    { id:'rules',   label:'KI-Regeln',  icon:'ti-robot'      },
    { id:'export',  label:'Export',     icon:'ti-download'   },
    { id:'data',    label:'Daten',      icon:'ti-database'   },
    { id:'import', label:'Bank Import', icon:'ti-building-bank' },
  ]

  return (
    <div className="fade-up">
      <div style={{ display:'flex', gap:6, marginBottom:'1.25rem', flexWrap:'wrap' }}>
        {SECTIONS.map(s=>(
          <button key={s.id} onClick={()=>setSection(s.id)} style={{
            display:'flex', alignItems:'center', gap:6, padding:'7px 14px',
            borderRadius:'var(--r-md)', fontSize:13, fontWeight:500, cursor:'pointer',
            border:'1px solid var(--brd)',
            background: section===s.id ? 'var(--violet-bg)' : 'var(--card2)',
            color:       section===s.id ? 'var(--violet)'    : 'var(--t2)',
          }}>
            <i className={`ti ${s.icon}`} />{s.label}
          </button>
        ))}
      </div>

      {section==='cats' && (
        <div>
          {cats.map(cat => {
            const info = getCatInfo(cat.name, cats)
            return (
              <Card key={cat.name} style={{ display:'flex', alignItems:'center', gap:12 }}>
                <div style={{ width:34, height:34, borderRadius:'50%', background:info.bg, color:info.c, display:'flex', alignItems:'center', justifyContent:'center', fontSize:16, flexShrink:0 }}>
                  <i className={`ti ${cat.icon}`} />
                </div>
                <div style={{ flex:1, fontSize:14, fontWeight:500, color:'var(--t1)' }}>{cat.name}</div>
                <button onClick={()=>delCat(cat.name)} style={{ color:'var(--t3)', background:'none', border:'none', cursor:'pointer', fontSize:18 }}>
                  <i className="ti ti-trash" />
                </button>
              </Card>
            )
          })}
          <Card style={{ background:'var(--card2)' }}>
            <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:12 }}>Neue Kategorie</div>
            <Field label="Name" style={{ marginBottom:10 }}>
              <Input value={newCatName} onChange={e=>setNewCatName(e.target.value)} placeholder="z.B. Sport" />
            </Field>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:12 }}>
              <Field label="Icon">
                <Sel value={newCatIcon} onChange={e=>setNewCatIcon(e.target.value)}>
                  {ICON_OPTS.map(i=><option key={i} value={i}>{i.replace('ti-','')}</option>)}
                </Sel>
              </Field>
              <Field label="Farbe">
                <Sel value={newCatColor} onChange={e=>setNewCatColor(e.target.value)}>
                  {COLOR_OPTS.map(c=><option key={c} value={c}>{c}</option>)}
                </Sel>
              </Field>
            </div>
            <Btn onClick={addCat}><i className="ti ti-plus" /> Hinzufügen</Btn>
          </Card>
        </div>
      )}

      {section==='sources' && (
        <div>
          {sources.map(s=>(
            <Card key={s} style={{ display:'flex', alignItems:'center', gap:12 }}>
              <i className="ti ti-credit-card" style={{ color:'var(--sky)', fontSize:20 }} />
              <div style={{ flex:1, fontSize:14, fontWeight:500, color:'var(--t1)' }}>{s}</div>
              <button onClick={()=>{ setSources(p=>p.filter(x=>x!==s)); toast('Quelle gelöscht') }} style={{ color:'var(--t3)', background:'none', border:'none', cursor:'pointer', fontSize:18 }}>
                <i className="ti ti-trash" />
              </button>
            </Card>
          ))}
          <Card style={{ background:'var(--card2)' }}>
            <div style={{ fontSize:13, fontWeight:600, color:'var(--t2)', marginBottom:12 }}>Neue Quelle</div>
            <div style={{ display:'flex', gap:8 }}>
              <Input value={newSource} onChange={e=>setNewSource(e.target.value)} placeholder="z.B. Sparkasse" style={{ flex:1 }} />
              <Btn onClick={addSource}><i className="ti ti-plus" /> Hinzufügen</Btn>
            </div>
          </Card>
        </div>
      )}

      {section==='rules' && (
        <div>
          <p style={{ fontSize:13, color:'var(--t2)', marginBottom:'1rem', lineHeight:1.5 }}>
            Schlüsselwörter pro Kategorie (kommagetrennt). Die KI nutzt diese zur automatischen Erkennung.
          </p>
          {cats.map(cat=>(
            <Card key={cat.name}>
              <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:8 }}>
                <CatIcon cat={cat.name} cats={cats} size={28} />
                <div style={{ fontSize:14, fontWeight:600, color:'var(--t1)' }}>{cat.name}</div>
              </div>
              <Input
                value={editRules[cat.name]||''}
                onChange={e=>setEditRules(p=>({ ...p, [cat.name]:e.target.value }))}
                placeholder="rewe, lidl, aldi, …"
                style={{ fontSize:12 }}
              />
            </Card>
          ))}
          <Btn onClick={saveRules} style={{ width:'100%', justifyContent:'center' }}>
            <i className="ti ti-check" /> KI-Regeln speichern
          </Btn>
        </div>
      )}
{section==='import' && (
  <SparkasseImport
    setTxs={setTxs} nid={nid} setNid={setNid}
    toast={toast} onBack={()=>setSection('cats')}
    cats={cats} rules={rules} existingTxs={txs}
  />
)}
      {section==='export' && (
        <div>
          {[
            { icon:'ti-file-spreadsheet', label:'Excel exportieren', sub:'.xlsx · 2 Tabellenblätter', c:'var(--mint)', bg:'var(--mint-bg)', fn:exportExcel },
            { icon:'ti-file-text',        label:'CSV exportieren',   sub:'Universelles Format',        c:'var(--sky)',  bg:'var(--sky-bg)',  fn:exportCSV  },
          ].map(o=>(
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
      )}

      {section==='data' && (
        <div>
          <Card style={{ border:'1px solid var(--coral-bg)' }}>
            <div style={{ fontSize:15, fontWeight:600, color:'var(--coral)', marginBottom:6 }}>Alle Daten zurücksetzen</div>
            <div style={{ fontSize:13, color:'var(--t2)', marginBottom:12, lineHeight:1.5 }}>Löscht alle Buchungen, Budgets und Einstellungen. Nicht rückgängig machbar.</div>
            <Btn variant="danger" onClick={resetAll}><i className="ti ti-trash" /> Zurücksetzen</Btn>
          </Card>
        </div>
      )}
    </div>
  )
}

// ── Import (CSV) ───────────────────────────────────────────────────────────

export function ImportView({ setTxs, nid, setNid, toast, onBack, cats, rules }) {
  const catNames = cats.map(c=>c.name)
  const [file, setFile]       = useState(null)
  const [headers, setHeaders] = useState([])
  const [preview, setPreview] = useState([])
  const [rows, setRows]       = useState([])
  const [map, setMap]         = useState({ date:'', desc:'', amount:'' })
  const [importing, setImporting] = useState(false)

  const onFile = e => {
    const f = e.target.files?.[0]; if (!f) return
    setFile(f)
    const reader = new FileReader()
    reader.onload = ev => {
      const wb  = XLSX.read(ev.target.result, { type:'binary' })
      const ws  = wb.Sheets[wb.SheetNames[0]]
      const data = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' })
      if (data.length < 2) { toast('Datei zu leer','err'); return }
      setHeaders(data[0]); setRows(data.slice(1)); setPreview(data.slice(1,6))
    }
    reader.readAsBinaryString(f)
  }

  const doImport = async () => {
    if (!map.date||!map.desc||!map.amount) { toast('Spalten zuordnen','err'); return }
    setImporting(true)
    const di=headers.indexOf(map.date), ni=headers.indexOf(map.desc), ai=headers.indexOf(map.amount)
    let id=nid; const newTxs=[]
    for (const row of rows) {
      const rawAmt = parseFloat(String(row[ai]).replace(',','.'))
      if (isNaN(rawAmt)) continue
      const rawDate = row[di]
      let date = today()
      if (typeof rawDate==='number') {
        const d = new Date(Math.round((rawDate-25569)*86400*1000))
        date = d.toISOString().split('T')[0]
      } else if (rawDate) {
        const parts = String(rawDate).split(/[.\-\/]/)
        if (parts.length===3) {
          date = parts[2].length===4
            ? `${parts[2]}-${parts[1].padStart(2,'0')}-${parts[0].padStart(2,'0')}`
            : `${parts[0]}-${parts[1].padStart(2,'0')}-${parts[2].padStart(2,'0')}`
        }
      }
      const desc = String(row[ni]||'Import')
      const cat  = aiCat(desc, rules)
      newTxs.push({ id:id++, desc, amount:Math.abs(rawAmt), type:rawAmt<0?'expense':'income', cat, date, source:'Import' })
    }
    setTxs(p=>[...newTxs,...p]); setNid(id)
    setImporting(false); toast(`${newTxs.length} Buchungen importiert ✓`); onBack()
  }

  return (
    <div className="fade-up">
      <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:'1.25rem' }}>
        <button onClick={onBack} style={{ color:'var(--t2)', fontSize:18 }} aria-label="Zurück">
          <i className="ti ti-arrow-left" />
        </button>
        <span style={{ fontSize:18, fontWeight:600 }}>CSV / Excel importieren</span>
      </div>

      {!file ? (
        <Card>
          <div style={{ textAlign:'center', padding:'1.5rem 0' }}>
            <i className="ti ti-upload" style={{ fontSize:36, color:'var(--t3)', display:'block', marginBottom:12 }} />
            <div style={{ fontSize:14, color:'var(--t2)', marginBottom:16 }}>CSV oder Excel-Datei von deiner Bank</div>
            <label style={{ display:'inline-flex', alignItems:'center', gap:8, padding:'10px 20px', borderRadius:'var(--r-md)', background:'var(--t1)', color:'var(--page)', fontSize:14, fontWeight:500, cursor:'pointer' }}>
              <i className="ti ti-file-upload" /> Datei wählen
              <input type="file" accept=".csv,.xlsx,.xls" onChange={onFile} style={{ display:'none' }} />
            </label>
          </div>
        </Card>
      ) : (
        <>
          <p style={{ fontSize:13, color:'var(--t2)', marginBottom:'1rem' }}>Welche Spalte enthält was?</p>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr 1fr', gap:8, marginBottom:'1rem' }}>
            {(['date','desc','amount']).map(k => (
              <Field key={k} label={k==='date'?'Datum':k==='desc'?'Beschreibung':'Betrag'}>
                <Sel value={map[k]} onChange={e=>setMap(p=>({ ...p, [k]:e.target.value }))}>
                  <option value="">– wählen –</option>
                  {headers.map(h=><option key={h} value={h}>{h}</option>)}
                </Sel>
              </Field>
            ))}
          </div>
          {preview.length>0 && (
            <div style={{ overflowX:'auto', marginBottom:'1rem' }}>
              <table style={{ fontSize:11, borderCollapse:'collapse', width:'100%' }}>
                <thead><tr>{headers.map(h=><th key={h} style={{ textAlign:'left', padding:'4px 8px', color:'var(--t3)', borderBottom:'1px solid var(--brd)' }}>{h}</th>)}</tr></thead>
                <tbody>{preview.map((r,i)=><tr key={i}>{r.map((v,j)=><td key={j} style={{ padding:'4px 8px', borderBottom:'1px solid var(--brd)', color:'var(--t1)' }}>{v}</td>)}</tr>)}</tbody>
              </table>
            </div>
          )}
          <div style={{ display:'flex', gap:8 }}>
            <Btn onClick={doImport} disabled={importing}>{importing?'…Importiere':`${rows.length} Zeilen importieren`}</Btn>
            <Btn variant="ghost" onClick={onBack}>Abbrechen</Btn>
          </div>
        </>
      )}
    </div>
   )
  }
  // SparkasseImport.jsx — Automatischer Sparkasse CSV Importer
// Füge diese Komponente in App.jsx ein (vor der letzten Zeile)

// ── Sparkasse Import ───────────────────────────────────────────────────────

function SparkasseImport({ setTxs, nid, setNid, toast, onBack, cats, rules, existingTxs }) {
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

const SPARKASSE_RULES = {
  Lebensmittel: 'rewe,lidl,aldi,edeka,penny,netto,kaufland,dm,rossmann,bäcker,metzger,wasgau,backerei,backshop,drogerie,müller,mueller,globus,spar,tegut,real,norma,marktkauf',
  Wohnen:       'miete,nebenkosten,strom,gas,wasser,internet,telefon,hausgeld,1und1,1+1,unitymedia,vodafone,telekom,o2,ewe,stadtwerke,stadtwerk,verbandsgemeinde',
  Transport:    'monatskarte,db,bahn,öpnv,tankstelle,shell,aral,bp,uber,taxi,flixbus,lufthansa,minera,kraftstoff,benzin,parkgebühr,parkgebuehr,autohaus,hyundai,kfz',
  Freizeit:     'netflix,spotify,amazon prime,disney,kino,theater,konzert,steam,playstation,gym,fitnessstudio,fitness,sport,voswinkel,supercell,riot games,epic games,google play,app store,kinoheld,happypottery,pottery,restaurant,pizzeria,ristorante,gastro,diner,eiscafe,sushi,hikari,dong,asia',
  Gesundheit:   'apotheke,arzt,krankenhaus,optiker,zahnarzt,physiotherapie,sanitätshaus,rats-apotheke',
  Versicherung: 'versicherung,allianz,aok,tkk,barmer,huk,continentale,europa verbund,samsung pay,gewerkschaft,verdi',
  Gehalt:       'gehalt,lohn,freelance,honorar,bundeskasse,renten service,rente,pension,lohn  gehalt,rente pensionszahlung,gutschr. ueberw. dauerauftr,zuwendung',
  Sonstiges:    'amazon,topstep,consors,easycredit,teambank,paypal,etsy,microsoft,google,anthropic,claude',
}


