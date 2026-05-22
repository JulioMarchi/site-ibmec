import { useState, useEffect, useCallback, useRef, useMemo } from 'react'
import * as XLSX from 'xlsx'
import { db } from './lib/supabase'
import {
  abrevCurso, buildCodigo, normalize,
  DIAS, DIAS_FULL, DIAS_LABEL,
  todayDia, fmtDate, nowTime, nowTimeWithSec, isAoVivo,
  CURSO_MAP
} from './lib/helpers'

/* ── XLSX parsers ──────────────────────────────────────────────── */
function parseAlocacao(buf) {
  const wb = XLSX.read(buf, { type: 'array' })
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' })
}
function parseEventos(buf) {
  const wb = XLSX.read(buf, { type: 'array' })
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '', raw: false })
}

/* ── Icons ────────────────────────────────────────────────────── */
const Icon = {
  Home: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z" /><polyline points="9 22 9 12 15 12 15 22" /></svg>,
  Grid: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="3" width="7" height="7" /><rect x="14" y="3" width="7" height="7" /><rect x="14" y="14" width="7" height="7" /><rect x="3" y="14" width="7" height="7" /></svg>,
  Calendar: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" /><line x1="16" y1="2" x2="16" y2="6" /><line x1="8" y1="2" x2="8" y2="6" /><line x1="3" y1="10" x2="21" y2="10" /></svg>,
  Users: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M23 21v-2a4 4 0 00-3-3.87M16 3.13a4 4 0 010 7.75" /></svg>,
  Upload: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" /></svg>,
  Chart: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="20" x2="18" y2="10" /><line x1="12" y1="20" x2="12" y2="4" /><line x1="6" y1="20" x2="6" y2="14" /></svg>,
  Settings: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z" /></svg>,
  Logout: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4" /><polyline points="16 17 21 12 16 7" /><line x1="21" y1="12" x2="9" y2="12" /></svg>,
  Plus: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" /></svg>,
  Edit: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7" /><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z" /></svg>,
  Trash: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6" /><path d="M19 6l-1 14a2 2 0 01-2 2H8a2 2 0 01-2-2L5 6" /><path d="M10 11v6M14 11v6" /><path d="M9 6V4a1 1 0 011-1h4a1 1 0 011 1v2" /></svg>,
  Bell: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9" /><path d="M13.73 21a2 2 0 01-3.46 0" /></svg>,
  X: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="18" y1="6" x2="6" y2="18" /><line x1="6" y1="6" x2="18" y2="18" /></svg>,
  Check: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12" /></svg>,
  Eye: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z" /><circle cx="12" cy="12" r="3" /></svg>,
  EyeOff: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0112 20c-7 0-11-8-11-8a18.45 18.45 0 015.06-5.94M9.9 4.24A9.12 9.12 0 0112 4c7 0 11 8 11 8a18.5 18.5 0 01-2.16 3.19m-6.72-1.07a3 3 0 11-4.24-4.24" /><line x1="1" y1="1" x2="23" y2="23" /></svg>,
  Lock: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="3" y="11" width="18" height="11" rx="2" ry="2" /><path d="M7 11V7a5 5 0 0110 0v4" /></svg>,
  Search: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8" /><line x1="21" y1="21" x2="16.65" y2="16.65" /></svg>,
  Refresh: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="23 4 23 10 17 10" /><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10" /></svg>,
  User: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2" /><circle cx="12" cy="7" r="4" /></svg>,
  Menu: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="3" y1="12" x2="21" y2="12" /><line x1="3" y1="6" x2="21" y2="6" /><line x1="3" y1="18" x2="21" y2="18" /></svg>,
  Clock: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10" /><polyline points="12 6 12 12 16 14" /></svg>,
  Ban: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10" /><line x1="4.93" y1="4.93" x2="19.07" y2="19.07" /></svg>,
  Monitor: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="2" y="3" width="20" height="14" rx="2" ry="2" /><line x1="8" y1="21" x2="16" y2="21" /><line x1="12" y1="17" x2="12" y2="21" /></svg>,
  Code: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="16 18 22 12 16 6" /><polyline points="8 6 2 12 8 18" /></svg>,
  Copy: () => <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2" /><path d="M5 15H4a2 2 0 01-2-2V4a2 2 0 012-2h9a2 2 0 012 2v1" /></svg>,
}

/* ── Toast ────────────────────────────────────────────────────── */
function Toast({ msg, type, onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 3500); return () => clearTimeout(t) }, [])
  return <div className={`toast toast-${type}`}>{type === 'success' ? <Icon.Check /> : <Icon.X />}{msg}</div>
}

/* ── Login ────────────────────────────────────────────────────── */
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState('')
  const [pass, setPass] = useState('')
  const [show, setShow] = useState(false)
  const [loading, setLoading] = useState(false)
  const [err, setErr] = useState('')

  async function handleSubmit(e) {
    e.preventDefault(); setLoading(true); setErr('')
    try {
      const { data, error } = await db.auth.signInWithPassword({ email, password: pass })
      if (error) throw error
      onLogin(data.user)
    } catch (e) { setErr(e.message || 'Erro ao entrar') }
    setLoading(false)
  }

  return (
    <div style={{ minHeight: '100vh', background: `linear-gradient(135deg,var(--navy-dark) 0%,var(--navy) 60%,#003a8c 100%)`, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 20 }}>
      <div style={{ width: '100%', maxWidth: 420 }}>
        <div style={{ textAlign: 'center', marginBottom: 36 }}>
          <div style={{ display: 'inline-flex', alignItems: 'center', gap: 12, background: 'rgba(255,255,255,.1)', padding: '12px 28px', borderRadius: 14, marginBottom: 20 }}>
            <div style={{ width: 10, height: 10, borderRadius: '50%', background: 'var(--gold)' }} />
            <span style={{ fontFamily: 'Sora', fontWeight: 800, fontSize: 30, color: '#fff', letterSpacing: '-0.5px' }}>ibmec</span>
          </div>
          <h1 style={{ color: '#fff', fontSize: 20, fontWeight: 600, marginBottom: 6 }}>Gestão de Campus</h1>
          <p style={{ color: 'rgba(255,255,255,.55)', fontSize: 14 }}>Acesse sua conta para continuar</p>
        </div>
        <div style={{ background: 'rgba(255,255,255,.07)', backdropFilter: 'blur(12px)', borderRadius: 16, padding: 32, border: '1px solid rgba(255,255,255,.12)' }}>
          <form onSubmit={handleSubmit} style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
            <div>
              <label className="label" style={{ color: 'rgba(255,255,255,.7)' }}>E-mail</label>
              <input className="input" type="email" value={email} onChange={e => setEmail(e.target.value)} placeholder="seu@email.com" style={{ background: 'rgba(255,255,255,.08)', border: '1.5px solid rgba(255,255,255,.15)', color: '#fff' }} required />
            </div>
            <div>
              <label className="label" style={{ color: 'rgba(255,255,255,.7)' }}>Senha</label>
              <div style={{ position: 'relative' }}>
                <input className="input" type={show ? 'text' : 'password'} value={pass} onChange={e => setPass(e.target.value)} placeholder="••••••••" style={{ background: 'rgba(255,255,255,.08)', border: '1.5px solid rgba(255,255,255,.15)', color: '#fff', paddingRight: 42 }} required />
                <button type="button" onClick={() => setShow(!show)} style={{ position: 'absolute', right: 12, top: '50%', transform: 'translateY(-50%)', background: 'none', border: 'none', cursor: 'pointer', color: 'rgba(255,255,255,.5)', display: 'flex' }}>{show ? <Icon.EyeOff /> : <Icon.Eye />}</button>
              </div>
            </div>
            {err && <div style={{ background: 'rgba(220,38,38,.15)', border: '1px solid rgba(220,38,38,.3)', color: '#fca5a5', padding: '10px 14px', borderRadius: 8, fontSize: 13 }}>{err}</div>}
            <button type="submit" className="btn btn-gold" style={{ width: '100%', justifyContent: 'center', padding: '12px', fontSize: 15, marginTop: 4 }} disabled={loading}>
              {loading ? <span className="loading-spin">↻</span> : 'Entrar'}
            </button>
          </form>
        </div>
        <p style={{ textAlign: 'center', marginTop: 14, fontSize: 13 }}>
          <a href="?modo=totem" style={{ color: 'rgba(255,255,255,.4)', textDecoration: 'none' }} target="_blank">🖥 Modo Totem (display público)</a>
        </p>
        <p style={{ textAlign: 'center', color: 'rgba(255,255,255,.3)', fontSize: 12, marginTop: 10 }}>Ibmec Campus System © {new Date().getFullYear()}</p>
      </div>
    </div>
  )
}

/* ── Sidebar ──────────────────────────────────────────────────── */
function Sidebar({ page, setPage, isAdmin, onLogout, collapsed }) {
  const items = [
    { id: 'dashboard', label: 'Dashboard', icon: 'Home' },
    { id: 'salas', label: 'Rotina do Campus', icon: 'Grid' },
    { id: 'grade', label: 'Grade Semanal', icon: 'Calendar' },
    { id: 'eventos', label: 'Eventos', icon: 'Bell' },
    ...(isAdmin ? [
      { id: 'planilhas', label: 'Planilhas', icon: 'Upload' },
      { id: 'cancelamentos', label: 'Cancelamentos', icon: 'Ban' },
      { id: 'usuarios', label: 'Usuários', icon: 'Users' },
      { id: 'relatorios', label: 'Relatórios', icon: 'Chart' },
    ] : []),
    { id: 'configuracoes', label: 'Configurações', icon: 'Settings' },
  ]
  const w = collapsed ? 64 : 220
  return (
    <div style={{ width: w, background: 'var(--navy)', display: 'flex', flexDirection: 'column', height: '100vh', position: 'fixed', left: 0, top: 0, zIndex: 100, transition: 'width .2s', flexShrink: 0 }}>
      <div style={{ padding: '20px 14px', borderBottom: '1px solid rgba(255,255,255,.1)', minHeight: 64, display: 'flex', alignItems: 'center' }}>
        {!collapsed && <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 8, height: 8, borderRadius: '50%', background: 'var(--gold)', flexShrink: 0 }} />
          <span style={{ fontFamily: 'Sora', fontWeight: 800, fontSize: 22, color: '#fff' }}>ibmec</span>
        </div>}
        {collapsed && <div style={{ width: 8, height: 8, borderRadius: '50%', background: 'var(--gold)', margin: '0 auto' }} />}
      </div>
      <nav style={{ flex: 1, padding: '12px 8px', display: 'flex', flexDirection: 'column', gap: 2, overflowY: 'auto' }}>
        {items.map(i => {
          const Ic = Icon[i.icon]
          return (
            <div key={i.id} className={`sidebar-item${page === i.id ? ' active' : ''}`} onClick={() => setPage(i.id)} title={i.label}>
              <Ic />{!collapsed && <span>{i.label}</span>}
            </div>
          )
        })}
      </nav>
      <div style={{ padding: '12px 8px', borderTop: '1px solid rgba(255,255,255,.1)' }}>
        <div className="sidebar-item" onClick={() => window.open('?modo=totem', '_blank')} title="Modo Totem">
          <Icon.Monitor />{!collapsed && <span>Modo Totem</span>}
        </div>
        <div className="sidebar-item" onClick={onLogout} title="Sair">
          <Icon.Logout />{!collapsed && <span>Sair</span>}
        </div>
      </div>
    </div>
  )
}

/* ── Dashboard ────────────────────────────────────────────────── */
function Dashboard({ aulas, eventos, cancelamentos }) {
  const [clock, setClock] = useState(nowTimeWithSec())
  useEffect(() => { const t = setInterval(() => setClock(nowTimeWithSec()), 1000); return () => clearInterval(t) }, [])

  const hoje = todayDia()
  const aulasHoje = aulas.filter(a => normalize(a['DIA DA SEMANA']) === normalize(hoje))
  const salasUnicas = [...new Set(aulas.map(a => a['SALA']).filter(Boolean))]
  const profs = [...new Set(aulas.map(a => a['PROFESSOR']).filter(Boolean))]
  const aulasAoVivo = aulasHoje.filter(a => isAoVivo(a['HORÁRIO']))
  const proxEventos = eventos.slice().sort((a, b) => new Date(a['Data'] || a['data'] || 0) - new Date(b['Data'] || b['data'] || 0))
    .filter(e => { try { return new Date(e['Data'] || e['data']) >= new Date() } catch { return true } }).slice(0, 5)

  const agora = new Date()
  const dateStr = agora.toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' })

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 24, flexWrap: 'wrap', gap: 12 }}>
        <div>
          <h2 style={{ fontSize: 22, color: 'var(--navy)', marginBottom: 4 }}>Dashboard</h2>
          <p style={{ fontSize: 13, color: 'var(--gray-400)', textTransform: 'capitalize' }}>{dateStr}</p>
        </div>
        <div style={{ background: 'var(--navy)', color: '#fff', borderRadius: 10, padding: '10px 20px', display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Sora' }}>
          <Icon.Clock />
          <span style={{ fontSize: 22, fontWeight: 700, letterSpacing: 2 }}>{clock}</span>
        </div>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(160px,1fr))', gap: 14, marginBottom: 24 }}>
        {[
          { num: salasUnicas.length, label: 'Salas Cadastradas', color: 'var(--navy)' },
          { num: aulasHoje.length, label: 'Aulas Hoje', color: 'var(--navy)' },
          { num: aulasAoVivo.length, label: 'Aulas ao Vivo', color: 'var(--success)' },
          { num: profs.length, label: 'Professores', color: 'var(--navy)' },
          { num: eventos.length, label: 'Eventos', color: 'var(--gold)' },
          { num: cancelamentos.length, label: 'Cancelamentos', color: 'var(--danger)' },
        ].map((s, i) => (
          <div key={i} className="stat-box">
            <div className="stat-num" style={{ color: s.color }}>{s.num}</div>
            <div className="stat-label">{s.label}</div>
          </div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <div className="card">
          <h3 style={{ fontSize: 15, marginBottom: 14, color: 'var(--navy)', display: 'flex', alignItems: 'center', gap: 6 }}><Icon.Clock />Aulas ao Vivo Agora</h3>
          {aulasAoVivo.length === 0 ? <p style={{ color: 'var(--gray-400)', fontSize: 14 }}>Nenhuma aula em andamento</p> :
            aulasAoVivo.slice(0, 8).map((a, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: '1px solid var(--gray-100)' }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--navy)' }}>{a['SALA']}</div>
                  <div style={{ fontSize: 12, color: 'var(--gray-600)' }}>{a['DISCIPLINA']}</div>
                  <div style={{ fontSize: 11, color: 'var(--gray-400)' }}>{a['PROFESSOR']}</div>
                </div>
                <span className="badge badge-success">● Ao vivo</span>
              </div>
            ))}
        </div>
        <div className="card">
          <h3 style={{ fontSize: 15, marginBottom: 14, color: 'var(--navy)', display: 'flex', alignItems: 'center', gap: 6 }}><Icon.Bell />Próximos Eventos</h3>
          {proxEventos.length === 0 ? <p style={{ color: 'var(--gray-400)', fontSize: 14 }}>Nenhum evento cadastrado</p> :
            proxEventos.map((e, i) => (
              <div key={i} style={{ padding: '8px 0', borderBottom: '1px solid var(--gray-100)' }}>
                <div style={{ fontSize: 13, fontWeight: 600 }}>{e['Nome do Evento'] || e['nome'] || 'Evento'}</div>
                <div style={{ fontSize: 12, color: 'var(--gray-400)' }}>{fmtDate(e['Data'] || e['data'])} · {e['Sala'] || e['sala'] || ''} · {e['Horário Inicio'] || e['horario_inicio'] || ''}</div>
              </div>
            ))}
        </div>
      </div>
      {aulasHoje.length > 0 && (
        <div className="card" style={{ marginTop: 16 }}>
          <h3 style={{ fontSize: 15, marginBottom: 14, color: 'var(--navy)' }}>Resumo de Hoje — {DIAS_FULL[hoje]}</h3>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8 }}>
            {[...new Set(aulasHoje.map(a => a['SALA']).filter(Boolean))].sort().map(sala => {
              const n = aulasHoje.filter(a => a['SALA'] === sala).length
              return <div key={sala} style={{ background: 'var(--gray-50)', border: '1px solid var(--gray-200)', borderRadius: 8, padding: '6px 12px', fontSize: 13 }}>
                <strong style={{ color: 'var(--navy)' }}>{sala}</strong><span style={{ color: 'var(--gray-400)', marginLeft: 6 }}>{n} aula{n !== 1 ? 's' : ''}</span>
              </div>
            })}
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Salas Hoje ───────────────────────────────────────────────── */
function SalasHoje({ aulas, eventos, cancelamentos }) {
  const [dia, setDia] = useState(todayDia())
  const [busca, setBusca] = useState('')
  const [view, setView] = useState('sala')

  const isCancelado = (row, tipo) => cancelamentos.some(c =>
    normalize(c.sala) === normalize(row['SALA']) &&
    normalize(c.horario) === normalize(row['HORÁRIO']) &&
    normalize(c.dia) === normalize(row['DIA DA SEMANA']) &&
    (tipo ? c.tipo === tipo : true)
  )

  const aulasDia = useMemo(() => {
    const map = new Map()
    aulas.filter(a => normalize(a['DIA DA SEMANA']) === normalize(dia)).forEach(row => {
      const key = `${normalize(row['SALA'])}|${normalize(row['HORÁRIO'])}|${normalize(row['DISCIPLINA'])}`
      if (!map.has(key)) map.set(key, { ...row, _codigos: [buildCodigo(row)], _cancelado: isCancelado(row, 'semana'), _canceladoDef: isCancelado(row, 'definitivo') })
      else { const ex = map.get(key); const c = buildCodigo(row); if (!ex._codigos.includes(c)) ex._codigos.push(c) }
    })
    return [...map.values()].filter(a => {
      if (!busca) return true
      const b = busca.toLowerCase()
      return (a['SALA'] || '').toLowerCase().includes(b) || (a['DISCIPLINA'] || '').toLowerCase().includes(b) || (a['PROFESSOR'] || '').toLowerCase().includes(b) || a._codigos.join(' ').toLowerCase().includes(b)
    })
  }, [aulas, dia, busca, cancelamentos])

  const bySala = useMemo(() => {
    const m = new Map()
    aulasDia.forEach(a => { const s = a['SALA'] || 'Sem sala'; if (!m.has(s)) m.set(s, []); m.get(s).push(a) })
    return m
  }, [aulasDia])

  const eventosDia = useMemo(() => eventos.filter(e => {
    const dds = normalize(e['Dia da Semana'] || e['dia_semana'] || '')
    return dds === normalize(dia) || dds === normalize(DIAS_LABEL[dia] || '')
  }), [eventos, dia])

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20, flexWrap: 'wrap', gap: 12 }}>
        <h2 style={{ fontSize: 22, color: 'var(--navy)' }}>Rotina do Campus</h2>
        <div style={{ display: 'flex', gap: 8 }}>
          <div style={{ position: 'relative' }}>
            <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--gray-400)', display: 'flex', width: 16, height: 16 }}><Icon.Search /></span>
            <input className="input" value={busca} onChange={e => setBusca(e.target.value)} placeholder="Buscar sala, disciplina..." style={{ paddingLeft: 36, width: 220 }} />
          </div>
          <button className="btn btn-ghost btn-sm" onClick={() => setView(view === 'sala' ? 'lista' : 'sala')}>{view === 'sala' ? 'Ver lista' : 'Ver por sala'}</button>
        </div>
      </div>
      <div className="dia-tabs">
        {DIAS.map(d => <div key={d} className={`dia-tab${dia === d ? ' active' : ''}`} onClick={() => setDia(d)}>{DIAS_FULL[d]}</div>)}
      </div>
      {eventosDia.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <h3 style={{ fontSize: 13, fontWeight: 700, color: 'var(--gray-600)', marginBottom: 10, textTransform: 'uppercase', letterSpacing: .5 }}>Eventos do Dia</h3>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
            {eventosDia.map((ev, i) => (
              <div key={i} style={{ background: 'var(--gold-pale)', border: '1.5px solid var(--gold)', borderRadius: 10, padding: '10px 14px', minWidth: 200 }}>
                <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--navy)' }}>{ev['Nome do Evento'] || ev['nome'] || 'Evento'}</div>
                <div style={{ fontSize: 12, color: 'var(--gray-600)', marginTop: 2 }}>{ev['Sala'] || ev['sala'] || ''} · {ev['Horário Inicio'] || ev['horario_inicio'] || ''} – {ev['Horario Termino'] || ev['horario_termino'] || ''}</div>
                {(ev['Responsável Reserva'] || ev['responsavel']) && <div style={{ fontSize: 12, color: 'var(--gray-400)' }}>{ev['Responsável Reserva'] || ev['responsavel']}</div>}
              </div>
            ))}
          </div>
        </div>
      )}
      {view === 'sala' ? (
        <div className="grid-salas">
          {[...bySala.entries()].sort((a, b) => a[0] > b[0] ? 1 : -1).map(([sala, list]) => (
            <div key={sala} className="sala-card">
              <div className="sala-header">
                <span style={{ fontFamily: 'Sora', fontWeight: 700, fontSize: 16 }}>{sala}</span>
                <span className="badge" style={{ background: 'rgba(255,255,255,.2)', color: '#fff', fontSize: 11 }}>{list.length} aula{list.length !== 1 ? 's' : ''}</span>
              </div>
              <div className="sala-body">
                {list.sort((a, b) => (a['HORÁRIO'] || '') > (b['HORÁRIO'] || '') ? 1 : -1).map((a, i) => {
                  const cls = a._canceladoDef ? 'cancelado-def' : a._cancelado ? 'cancelado' : ''
                  const live = isAoVivo(a['HORÁRIO'])
                  return (
                    <div key={i} className={`aula-item ${cls}`} style={{ borderLeftColor: live && !a._cancelado && !a._canceladoDef ? 'var(--success)' : undefined }}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 8, marginBottom: 4 }}>
                        <span style={{ fontSize: 12, fontWeight: 700, color: 'var(--navy)' }}>{a['HORÁRIO']}</span>
                        <div style={{ display: 'flex', gap: 4 }}>
                          {live && !a._cancelado && !a._canceladoDef && <span className="badge badge-success" style={{ fontSize: 10 }}>● Ao vivo</span>}
                          {a._cancelado && !a._canceladoDef && <span className="badge badge-warning" style={{ fontSize: 10 }}>Cancelado</span>}
                          {a._canceladoDef && <span className="badge badge-danger" style={{ fontSize: 10 }}>Canc. Def.</span>}
                        </div>
                      </div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: 'var(--gray-800)', marginBottom: 2 }}>{a['DISCIPLINA']}</div>
                      <div style={{ fontSize: 12, color: 'var(--gray-600)', marginBottom: 6 }}>{a['PROFESSOR']}</div>
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                        {a._codigos.map((c, ci) => <span key={ci} className="tag-turma">{c}</span>)}
                      </div>
                    </div>
                  )
                })}
              </div>
            </div>
          ))}
          {bySala.size === 0 && <div style={{ gridColumn: '1/-1', textAlign: 'center', padding: 60, color: 'var(--gray-400)' }}>Nenhuma aula encontrada para {DIAS_FULL[dia]}</div>}
        </div>
      ) : (
        <div className="card">
          <div className="table-container">
            <table>
              <thead><tr><th>Sala</th><th>Horário</th><th>Disciplina</th><th>Professor</th><th>Turmas</th><th>Status</th></tr></thead>
              <tbody>
                {aulasDia.sort((a, b) => (a['HORÁRIO'] || '') > (b['HORÁRIO'] || '') ? 1 : -1).map((a, i) => (
                  <tr key={i}>
                    <td style={{ fontWeight: 600 }}>{a['SALA']}</td>
                    <td><span className="badge badge-navy">{a['HORÁRIO']}</span></td>
                    <td>{a['DISCIPLINA']}</td>
                    <td style={{ color: 'var(--gray-600)' }}>{a['PROFESSOR']}</td>
                    <td><div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>{a._codigos.map((c, ci) => <span key={ci} className="tag-turma">{c}</span>)}</div></td>
                    <td>
                      {a._canceladoDef ? <span className="badge badge-danger">Canc. Def.</span> :
                        a._cancelado ? <span className="badge badge-warning">Cancelado</span> :
                          isAoVivo(a['HORÁRIO']) ? <span className="badge badge-success">● Ao vivo</span> :
                            <span className="badge badge-gray">Normal</span>}
                    </td>
                  </tr>
                ))}
                {aulasDia.length === 0 && <tr><td colSpan={6} style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}>Nenhuma aula encontrada</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Grade Semanal ────────────────────────────────────────────── */
function GradeSemanal({ aulas, cancelamentos }) {
  const [sala, setSala] = useState('')
  const [prof, setProf] = useState('')
  const salas = [...new Set(aulas.map(a => a['SALA']).filter(Boolean))].sort()
  const profs = [...new Set(aulas.map(a => a['PROFESSOR']).filter(Boolean))].sort()

  const isCancelado = row => cancelamentos.some(c =>
    normalize(c.sala) === normalize(row['SALA']) &&
    normalize(c.horario) === normalize(row['HORÁRIO']) &&
    normalize(c.dia) === normalize(row['DIA DA SEMANA'])
  )

  const grade = useMemo(() => {
    const map = new Map()
    aulas.filter(a => (!sala || normalize(a['SALA']) === normalize(sala)) && (!prof || normalize(a['PROFESSOR']) === normalize(prof))).forEach(row => {
      const key = `${normalize(row['SALA'])}|${normalize(row['DIA DA SEMANA'])}|${normalize(row['HORÁRIO'])}|${normalize(row['DISCIPLINA'])}`
      if (!map.has(key)) map.set(key, { ...row, _codigos: [buildCodigo(row)], _cancelado: isCancelado(row) })
      else map.get(key)._codigos.push(buildCodigo(row))
    })
    return [...map.values()]
  }, [aulas, sala, prof, cancelamentos])

  const horarios = [...new Set(grade.map(a => a['HORÁRIO']).filter(Boolean))].sort()
  const diasUsados = DIAS.filter(d => grade.some(a => normalize(a['DIA DA SEMANA']) === normalize(d)))
  const hoje = todayDia()

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20, flexWrap: 'wrap', gap: 12 }}>
        <h2 style={{ fontSize: 22, color: 'var(--navy)' }}>Grade Semanal</h2>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <select className="input" value={sala} onChange={e => setSala(e.target.value)} style={{ width: 'auto', minWidth: 160 }}>
            <option value="">Todas as salas</option>
            {salas.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <select className="input" value={prof} onChange={e => setProf(e.target.value)} style={{ width: 'auto', minWidth: 180 }}>
            <option value="">Todos os professores</option>
            {profs.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </div>
      </div>
      <div style={{ overflowX: 'auto' }}>
        <table style={{ minWidth: 700 }}>
          <thead>
            <tr>
              <th style={{ minWidth: 80 }}>Horário</th>
              {diasUsados.map(d => <th key={d} style={{ minWidth: 130, background: d === hoje ? 'rgba(0,36,84,.06)' : undefined }}>{DIAS_FULL[d]}{d === hoje && <span className="badge badge-gold" style={{ marginLeft: 6, fontSize: 10 }}>Hoje</span>}</th>)}
            </tr>
          </thead>
          <tbody>
            {horarios.map(h => (
              <tr key={h}>
                <td style={{ fontWeight: 700, color: 'var(--navy)', verticalAlign: 'top', whiteSpace: 'nowrap' }}>{h}</td>
                {diasUsados.map(d => {
                  const items = grade.filter(a => normalize(a['HORÁRIO']) === normalize(h) && normalize(a['DIA DA SEMANA']) === normalize(d))
                  return (
                    <td key={d} style={{ verticalAlign: 'top', background: d === hoje ? 'rgba(0,36,84,.02)' : undefined }}>
                      {items.map((a, i) => (
                        <div key={i} style={{ marginBottom: 6, padding: '6px 8px', borderRadius: 6, background: a._cancelado ? '#fee2e2' : '#eff6ff', borderLeft: `3px solid ${a._cancelado ? 'var(--danger)' : 'var(--info)'}`, opacity: a._cancelado ? .5 : 1 }}>
                          {!sala && <div style={{ fontSize: 11, fontWeight: 700, color: 'var(--gray-400)' }}>{a['SALA']}</div>}
                          <div style={{ fontSize: 12, fontWeight: 600, color: 'var(--gray-800)', marginBottom: 1 }}>{a['DISCIPLINA']}</div>
                          <div style={{ fontSize: 11, color: 'var(--gray-600)', marginBottom: 3 }}>{a['PROFESSOR']}</div>
                          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 2 }}>
                            {a._codigos.map((c, ci) => <span key={ci} style={{ fontSize: 10, background: 'var(--navy)', color: '#fff', padding: '1px 5px', borderRadius: 4 }}>{c}</span>)}
                          </div>
                          {a._cancelado && <div style={{ fontSize: 10, color: 'var(--danger)', fontWeight: 700, marginTop: 2 }}>CANCELADO</div>}
                        </div>
                      ))}
                    </td>
                  )
                })}
              </tr>
            ))}
            {horarios.length === 0 && <tr><td colSpan={diasUsados.length + 1} style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}>Nenhuma aula encontrada</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  )
}

/* ── Eventos ──────────────────────────────────────────────────── */
function Eventos({ eventos, isAdmin, onAddEvento }) {
  const [busca, setBusca] = useState('')
  const filtered = eventos.filter(e => {
    const b = busca.toLowerCase()
    return !b || (e['Nome do Evento'] || e['nome'] || '').toLowerCase().includes(b) || (e['Sala'] || e['sala'] || '').toLowerCase().includes(b) || (e['Responsável Reserva'] || e['responsavel'] || '').toLowerCase().includes(b)
  })

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20, flexWrap: 'wrap', gap: 12 }}>
        <h2 style={{ fontSize: 22, color: 'var(--navy)' }}>Eventos</h2>
        <div style={{ display: 'flex', gap: 8 }}>
          <div style={{ position: 'relative' }}>
            <span style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--gray-400)', display: 'flex', width: 16, height: 16 }}><Icon.Search /></span>
            <input className="input" value={busca} onChange={e => setBusca(e.target.value)} placeholder="Buscar evento..." style={{ paddingLeft: 36, width: 200 }} />
          </div>
          {isAdmin && <button className="btn btn-gold" onClick={onAddEvento}><Icon.Plus />Novo Evento</button>}
        </div>
      </div>
      <div className="card">
        <div className="table-container">
          <table>
            <thead><tr><th>Data</th><th>Dia</th><th>Nome do Evento</th><th>Sala</th><th>Horário</th><th>Responsável</th><th>Observações</th></tr></thead>
            <tbody>
              {filtered.map((e, i) => (
                <tr key={i}>
                  <td style={{ whiteSpace: 'nowrap', fontWeight: 600, color: 'var(--navy)' }}>{fmtDate(e['Data'] || e['data'])}</td>
                  <td style={{ whiteSpace: 'nowrap', color: 'var(--gray-600)' }}>{e['Dia da Semana'] || e['dia_semana'] || '-'}</td>
                  <td style={{ fontWeight: 500 }}>{e['Nome do Evento'] || e['nome'] || '-'}</td>
                  <td><span className="badge badge-navy">{e['Sala'] || e['sala'] || '-'}</span></td>
                  <td style={{ whiteSpace: 'nowrap' }}>{(e['Horário Inicio'] || e['horario_inicio'] || '')} {(e['Horario Termino'] || e['horario_termino']) ? '– ' + (e['Horario Termino'] || e['horario_termino']) : ''}</td>
                  <td style={{ color: 'var(--gray-600)' }}>{e['Responsável Reserva'] || e['responsavel'] || '-'}</td>
                  <td style={{ color: 'var(--gray-400)', fontSize: 12, maxWidth: 150 }}>{e['Observações'] || e['observacoes'] || '-'}</td>
                </tr>
              ))}
              {filtered.length === 0 && <tr><td colSpan={7} style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}>Nenhum evento encontrado</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  )
}

/* ── Cancelamentos Admin ──────────────────────────────────────── */
function CancelamentosAdmin({ toast }) {
  const [lista, setLista] = useState([])
  const [loading, setLoading] = useState(true)
  const [modal, setModal] = useState(false)
  const [alocacaoRows, setAlocacaoRows] = useState([])
  const [form, setForm] = useState({ sala: '', dia: '', horario: '', disciplina: '', tipo: 'semana', motivo: '' })

  async function load() {
    setLoading(true)
    const { data } = await db.from('cancelamentos').select('*').order('created_at', { ascending: false })
    setLista(data || []); setLoading(false)
  }
  async function loadAlocacao() {
    const { data } = await db.from('alocacao').select('sala,dia_semana,horario,disciplina').limit(1000)
    setAlocacaoRows(data || [])
  }
  useEffect(() => { load(); loadAlocacao() }, [])

  const salasDisp = [...new Set(alocacaoRows.map(r => r.sala).filter(Boolean))].sort()
  const diasDisp = [...new Set(alocacaoRows.filter(r => !form.sala || normalize(r.sala) === normalize(form.sala)).map(r => r.dia_semana).filter(Boolean))]
  const horariosDisp = [...new Set(alocacaoRows.filter(r =>
    (!form.sala || normalize(r.sala) === normalize(form.sala)) &&
    (!form.dia || normalize(r.dia_semana) === normalize(form.dia))
  ).map(r => r.horario).filter(Boolean))].sort()
  const disciplinasDisp = [...new Set(alocacaoRows.filter(r =>
    (!form.sala || normalize(r.sala) === normalize(form.sala)) &&
    (!form.dia || normalize(r.dia_semana) === normalize(form.dia)) &&
    (!form.horario || normalize(r.horario) === normalize(form.horario))
  ).map(r => r.disciplina).filter(Boolean))].sort()

  async function save() {
    if (!form.sala || !form.dia || !form.horario) { toast('Preencha sala, dia e horário', 'error'); return }
    const { error } = await db.from('cancelamentos').insert([form])
    if (!error) { toast('Cancelamento registrado!', 'success'); setModal(false); setForm({ sala: '', dia: '', horario: '', disciplina: '', tipo: 'semana', motivo: '' }); load() }
    else toast(error.message, 'error')
  }
  async function remove(id) {
    if (!confirm('Remover este cancelamento?')) return
    await db.from('cancelamentos').delete().eq('id', id)
    toast('Cancelamento removido', 'success'); load()
  }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, color: 'var(--navy)' }}>Cancelamentos</h2>
        <button className="btn btn-danger" onClick={() => setModal(true)}><Icon.Plus />Registrar Cancelamento</button>
      </div>
      <div className="card">
        {loading ? <p style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}><span className="loading-spin">↻</span> Carregando...</p> : (
          <div className="table-container">
            <table>
              <thead><tr><th>Sala</th><th>Dia</th><th>Horário</th><th>Disciplina</th><th>Tipo</th><th>Motivo</th><th>Registrado em</th><th>Ação</th></tr></thead>
              <tbody>
                {lista.map(c => (
                  <tr key={c.id}>
                    <td style={{ fontWeight: 600 }}>{c.sala}</td>
                    <td>{c.dia}</td>
                    <td><span className="badge badge-navy">{c.horario}</span></td>
                    <td style={{ color: 'var(--gray-600)' }}>{c.disciplina || '—'}</td>
                    <td>{c.tipo === 'definitivo' ? <span className="badge badge-danger">Definitivo</span> : <span className="badge badge-warning">Só esta semana</span>}</td>
                    <td style={{ color: 'var(--gray-400)', fontSize: 12, maxWidth: 160 }}>{c.motivo || '—'}</td>
                    <td style={{ color: 'var(--gray-400)', fontSize: 12 }}>{c.created_at ? new Date(c.created_at).toLocaleString('pt-BR') : '—'}</td>
                    <td><button className="btn btn-danger btn-sm" onClick={() => remove(c.id)}><Icon.Trash /></button></td>
                  </tr>
                ))}
                {lista.length === 0 && <tr><td colSpan={8} style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}>Nenhum cancelamento registrado</td></tr>}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {modal && (
        <div className="modal-overlay">
          <div className="modal">
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 20 }}>
              <h3 style={{ color: 'var(--navy)' }}>Registrar Cancelamento</h3>
              <button className="btn btn-ghost btn-sm" onClick={() => setModal(false)}><Icon.X /></button>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div><label className="label">Sala</label>
                <select className="input" value={form.sala} onChange={e => setForm({ ...form, sala: e.target.value, dia: '', horario: '', disciplina: '' })}>
                  <option value="">Selecione a sala...</option>
                  {salasDisp.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
              <div><label className="label">Dia da Semana</label>
                <select className="input" value={form.dia} onChange={e => setForm({ ...form, dia: e.target.value, horario: '', disciplina: '' })}>
                  <option value="">Selecione o dia...</option>
                  {diasDisp.map(d => <option key={d} value={d}>{DIAS_FULL[normalize(d)] || d}</option>)}
                </select>
              </div>
              <div><label className="label">Horário</label>
                <select className="input" value={form.horario} onChange={e => setForm({ ...form, horario: e.target.value, disciplina: '' })}>
                  <option value="">Selecione o horário...</option>
                  {horariosDisp.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div><label className="label">Disciplina (opcional)</label>
                <select className="input" value={form.disciplina} onChange={e => setForm({ ...form, disciplina: e.target.value })}>
                  <option value="">Todas neste horário/sala</option>
                  {disciplinasDisp.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>
              <div><label className="label">Tipo</label>
                <select className="input" value={form.tipo} onChange={e => setForm({ ...form, tipo: e.target.value })}>
                  <option value="semana">Cancelado (só esta semana)</option>
                  <option value="definitivo">Cancelado Definitivamente</option>
                </select>
              </div>
              <div><label className="label">Motivo (opcional)</label>
                <input className="input" value={form.motivo} onChange={e => setForm({ ...form, motivo: e.target.value })} placeholder="Ex: professor em viagem" />
              </div>
              <button className="btn btn-danger" style={{ marginTop: 8 }} onClick={save}>Confirmar Cancelamento</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Planilhas ────────────────────────────────────────────────── */
function Planilhas({ onAlocacaoUpload, onEventosUpload, toast }) {
  const alocRef = useRef()
  const evRef = useRef()
  const [loadingAloc, setLoadingAloc] = useState(false)
  const [loadingEv, setLoadingEv] = useState(false)
  const [lastAloc, setLastAloc] = useState(null)
  const [lastEv, setLastEv] = useState(null)

  useEffect(() => {
    db.from('alocacao').select('created_at').order('created_at', { ascending: false }).limit(1).then(({ data }) => { if (data && data[0]) setLastAloc(data[0].created_at) })
    db.from('eventos').select('criado_em').order('criado_em', { ascending: false }).limit(1).then(({ data }) => { if (data && data[0]) setLastEv(data[0].criado_em) })
  }, [])

  async function handleAloc(e) {
    const file = e.target.files[0]; if (!file) return
    setLoadingAloc(true)
    try {
      const buf = await file.arrayBuffer()
      const rows = parseAlocacao(new Uint8Array(buf))
      const mapped = rows.map(r => ({
        num_seq: String(r['NUM_SEQ_TURMA'] || ''), unidade: r['UNIDADE'] || '', curso: r['CURSO'] || '',
        periodo: String(r['PERÍODO'] || ''), turma: String(r['TURMA'] || ''), grupo: String(r['GRUPO'] || ''),
        codigo: String(r['CÓDIGO'] || ''), disciplina: r['DISCIPLINA'] || '', professor: r['PROFESSOR'] || '',
        dia_semana: r['DIA DA SEMANA'] || '', horario: r['HORÁRIO'] || '', sala: r['SALA'] || '',
        capacidade: String(r['CAPACIDADE'] || ''), raw: r,
      }))
      await db.from('alocacao').delete().gt('id', 0)
      const { error } = await db.from('alocacao').insert(mapped)
      if (!error) { toast(`Planilha importada! ${rows.length} linhas carregadas.`, 'success'); setLastAloc(new Date().toISOString()); onAlocacaoUpload(rows) }
      else toast('Erro ao importar: ' + error.message, 'error')
    } catch (err) { toast('Erro ao processar arquivo: ' + err.message, 'error') }
    setLoadingAloc(false); alocRef.current.value = ''
  }

  async function handleEv(e) {
    const file = e.target.files[0]; if (!file) return
    setLoadingEv(true)
    try {
      const buf = await file.arrayBuffer()
      const rows = parseEventos(new Uint8Array(buf))
      const mapped = rows.map(r => ({
        data: String(r['Data'] || r['DATA'] || ''),
        dia_semana: r['Dia da Semana'] || r['DIA DA SEMANA'] || '',
        nome: r['Nome do Evento'] || r['NOME DO EVENTO'] || '',
        responsavel: r['Responsável Reserva'] || '',
        sala: r['Sala'] || r['SALA'] || '',
        capacidade: String(r['Capacidade'] || r['CAPACIDADE'] || ''),
        horario_inicio: r['Horário Inicio'] || r['HORÁRIO INICIO'] || '',
        horario_termino: r['Horario Termino'] || r['HORÁRIO TÉRMINO'] || '',
        observacoes: r['Observações'] || r['OBSERVAÇÕES'] || '',
        assistente: r['Assistente/Supervisão'] || '',
      }))
      await db.from('eventos').delete().gt('id', 0)
      const { error } = await db.from('eventos').insert(mapped)
      if (!error) { toast(`Eventos importados! ${rows.length} eventos carregados.`, 'success'); setLastEv(new Date().toISOString()); onEventosUpload(rows) }
      else toast('Erro ao importar: ' + error.message, 'error')
    } catch (err) { toast('Erro ao processar arquivo: ' + err.message, 'error') }
    setLoadingEv(false); evRef.current.value = ''
  }

  return (
    <div>
      <h2 style={{ marginBottom: 20, fontSize: 22, color: 'var(--navy)' }}>Gerenciar Planilhas</h2>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 24 }}>
        <div className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 12 }}>
            <div>
              <h3 style={{ fontSize: 16, color: 'var(--navy)' }}>Planilha de Alocação de Salas</h3>
              <p style={{ fontSize: 13, color: 'var(--gray-400)', marginTop: 4 }}>Substitui a grade atual de aulas</p>
            </div>
            {lastAloc && <span style={{ fontSize: 11, color: 'var(--gray-400)' }}>Atualizada: {new Date(lastAloc).toLocaleString('pt-BR')}</span>}
          </div>
          <div className="upload-area" onClick={() => alocRef.current.click()}>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
              <span style={{ color: 'var(--navy)', display: 'block', width: 28, height: 28 }}><Icon.Upload /></span>
              <span style={{ fontSize: 14, color: 'var(--gray-600)' }}>Clique para selecionar arquivo</span>
              <span style={{ fontSize: 12, color: 'var(--gray-400)' }}>.xlsx · .xls</span>
            </div>
          </div>
          <input ref={alocRef} type="file" accept=".xlsx,.xls" onChange={handleAloc} style={{ display: 'none' }} />
          {loadingAloc && <p style={{ textAlign: 'center', marginTop: 10, color: 'var(--gray-400)', fontSize: 13 }}><span className="loading-spin">↻</span> Processando...</p>}
          <div style={{ marginTop: 12, padding: '10px 14px', background: 'var(--gray-50)', borderRadius: 8, fontSize: 12, color: 'var(--gray-600)' }}>
            <strong>Cabeçalho esperado:</strong> NUM_SEQ_TURMA, UNIDADE, CURSO, PERÍODO, TURMA, GRUPO, CÓDIGO, DISCIPLINA, PROFESSOR, DIA DA SEMANA, HORÁRIO, SALA, CAPACIDADE
          </div>
        </div>
        <div className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 12 }}>
            <div>
              <h3 style={{ fontSize: 16, color: 'var(--navy)' }}>Planilha de Eventos</h3>
              <p style={{ fontSize: 13, color: 'var(--gray-400)', marginTop: 4 }}>Atualizada automaticamente a cada hora</p>
            </div>
            {lastEv && <span style={{ fontSize: 11, color: 'var(--gray-400)' }}>Atualizada: {new Date(lastEv).toLocaleString('pt-BR')}</span>}
          </div>
          <div className="upload-area" onClick={() => evRef.current.click()}>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
              <span style={{ color: 'var(--gold)', display: 'block', width: 28, height: 28 }}><Icon.Upload /></span>
              <span style={{ fontSize: 14, color: 'var(--gray-600)' }}>Clique para selecionar arquivo</span>
              <span style={{ fontSize: 12, color: 'var(--gray-400)' }}>.xlsx · .xls</span>
            </div>
          </div>
          <input ref={evRef} type="file" accept=".xlsx,.xls" onChange={handleEv} style={{ display: 'none' }} />
          {loadingEv && <p style={{ textAlign: 'center', marginTop: 10, color: 'var(--gray-400)', fontSize: 13 }}><span className="loading-spin">↻</span> Processando...</p>}
          <div style={{ marginTop: 12, padding: '10px 14px', background: 'var(--gray-50)', borderRadius: 8, fontSize: 12, color: 'var(--gray-600)' }}>
            <strong>Cabeçalho esperado:</strong> Data, Dia da Semana, Nome do Evento, Responsável Reserva, Sala, Capacidade, Horário Inicio, Horario Termino, Observações
          </div>
        </div>
      </div>
      <div className="card" style={{ background: 'var(--navy)', color: '#fff', border: 'none' }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 12 }}>
          <div>
            <h3 style={{ fontSize: 15, marginBottom: 4 }}>Modo Totem / Display Público</h3>
            <p style={{ fontSize: 13, color: 'rgba(255,255,255,.6)' }}>Link para exibir a rotina do campus em telas públicas (TV, totem) sem necessidade de login.</p>
          </div>
          <button className="btn btn-gold" onClick={() => window.open('?modo=totem', '_blank')}><Icon.Monitor />Abrir Totem</button>
        </div>
      </div>
    </div>
  )
}

/* ── Usuários ─────────────────────────────────────────────────── */
function Usuarios({ toast, currentUser }) {
  const [users, setUsers] = useState([])
  const [loading, setLoading] = useState(true)
  const [modal, setModal] = useState(false)
  const [form, setForm] = useState({ email: '', password: '', nome: '', role: 'user' })
  const [saving, setSaving] = useState(false)

  async function loadUsers() {
    setLoading(true)
    const { data } = await db.from('user_profiles').select('*').order('created_at', { ascending: false })
    setUsers(data || []); setLoading(false)
  }
  useEffect(() => { loadUsers() }, [])

  async function createUser() {
    if (!form.email || !form.password) { toast('E-mail e senha são obrigatórios', 'error'); return }
    if (form.password.length < 6) { toast('Senha deve ter no mínimo 6 caracteres', 'error'); return }
    setSaving(true)
    const { data, error } = await db.auth.signUp({ email: form.email, password: form.password })
    if (error) { toast(error.message, 'error'); setSaving(false); return }
    await db.from('user_profiles').insert([{ id: data.user?.id, email: form.email, nome: form.nome, role: form.role }])
    toast('Usuário criado! Pode ser necessário confirmar o e-mail.', 'success')
    setModal(false); setForm({ email: '', password: '', nome: '', role: 'user' }); loadUsers(); setSaving(false)
  }

  async function changeRole(id, role) {
    await db.from('user_profiles').update({ role }).eq('id', id)
    toast('Papel atualizado!', 'success'); loadUsers()
  }

  async function deleteUser(id) {
    if (!confirm('Remover usuário?')) return
    await db.from('user_profiles').delete().eq('id', id)
    toast('Usuário removido', 'success'); loadUsers()
  }

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, color: 'var(--navy)' }}>Usuários</h2>
        <button className="btn btn-gold" onClick={() => setModal(true)}><Icon.Plus />Novo Usuário</button>
      </div>
      <div className="card" style={{ marginBottom: 12, background: '#fffbeb', border: '1px solid #fde68a', fontSize: 13, color: '#92400e', padding: '12px 16px' }}>
        ⚠️ Para que novos usuários não precisem confirmar e-mail, desative <strong>Email Confirmations</strong> em Authentication → Settings no painel do Supabase.
      </div>
      <div className="card">
        {loading ? <p style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}><span className="loading-spin">↻</span> Carregando...</p> : (
          <div className="table-container">
            <table>
              <thead><tr><th>Nome</th><th>E-mail</th><th>Papel</th><th>Criado em</th><th>Ações</th></tr></thead>
              <tbody>
                {users.map(u => (
                  <tr key={u.id}>
                    <td style={{ fontWeight: 500 }}>{u.nome || '—'}</td>
                    <td style={{ color: 'var(--gray-600)' }}>{u.email}</td>
                    <td>
                      <select className="input" style={{ width: 'auto' }} value={u.role || 'user'} onChange={e => changeRole(u.id, e.target.value)}>
                        <option value="user">Usuário</option>
                        <option value="sub-admin">Sub-Admin</option>
                        <option value="admin">Admin</option>
                      </select>
                    </td>
                    <td style={{ color: 'var(--gray-400)', fontSize: 12 }}>{u.created_at ? new Date(u.created_at).toLocaleDateString('pt-BR') : '-'}</td>
                    <td>
                      {u.id !== currentUser?.id && <button className="btn btn-danger btn-sm" onClick={() => deleteUser(u.id)}><Icon.Trash /></button>}
                      {u.id === currentUser?.id && <span className="badge badge-info">Você</span>}
                    </td>
                  </tr>
                ))}
                {users.length === 0 && <tr><td colSpan={5} style={{ textAlign: 'center', padding: 40, color: 'var(--gray-400)' }}>Nenhum usuário cadastrado</td></tr>}
              </tbody>
            </table>
          </div>
        )}
      </div>
      {modal && (
        <div className="modal-overlay">
          <div className="modal">
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 20 }}>
              <h3 style={{ color: 'var(--navy)' }}>Novo Usuário</h3>
              <button className="btn btn-ghost btn-sm" onClick={() => setModal(false)}><Icon.X /></button>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
              <div><label className="label">Nome</label><input className="input" value={form.nome} onChange={e => setForm({ ...form, nome: e.target.value })} placeholder="Nome completo" /></div>
              <div><label className="label">E-mail</label><input className="input" type="email" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} placeholder="email@ibmec.edu.br" /></div>
              <div><label className="label">Senha</label><input className="input" type="password" value={form.password} onChange={e => setForm({ ...form, password: e.target.value })} placeholder="Mínimo 6 caracteres" /></div>
              <div><label className="label">Papel</label>
                <select className="input" value={form.role} onChange={e => setForm({ ...form, role: e.target.value })}>
                  <option value="user">Usuário</option>
                  <option value="sub-admin">Sub-Admin</option>
                  <option value="admin">Admin</option>
                </select>
              </div>
              <button className="btn btn-primary" style={{ marginTop: 8 }} onClick={createUser} disabled={saving}>
                {saving ? <><span className="loading-spin">↻</span> Criando...</> : 'Criar Usuário'}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Relatórios ───────────────────────────────────────────────── */
function Relatorios({ aulas, eventos, cancelamentos }) {
  const [sqlModal, setSqlModal] = useState(false)
  const salasUnicas = [...new Set(aulas.map(a => a['SALA']).filter(Boolean))]
  const profs = [...new Set(aulas.map(a => a['PROFESSOR']).filter(Boolean))]
  const cursos = [...new Set(aulas.map(a => a['CURSO']).filter(Boolean))]
  const aulasPerDia = DIAS.map(d => ({ dia: DIAS_FULL[d], count: aulas.filter(a => normalize(a['DIA DA SEMANA']) === normalize(d)).length }))
  const top5Salas = salasUnicas.map(s => ({ sala: s, count: aulas.filter(a => a['SALA'] === s).length })).sort((a, b) => b.count - a.count).slice(0, 5)
  const maxS = Math.max(...top5Salas.map(s => s.count), 1)
  const maxD = Math.max(...aulasPerDia.map(d => d.count), 1)
  const cursosDist = Object.entries(CURSO_MAP).map(([nome, abrev]) => ({ abrev, count: aulas.filter(a => normalize(a['CURSO']) === normalize(nome)).length })).filter(c => c.count > 0).sort((a, b) => b.count - a.count)

  const SQL = `-- Execute este SQL no Editor SQL do Supabase
-- https://app.supabase.com → SQL Editor

-- 1. Tabela de perfis de usuário
CREATE TABLE IF NOT EXISTS public.user_profiles (
  id UUID REFERENCES auth.users ON DELETE CASCADE PRIMARY KEY,
  email TEXT, nome TEXT, role TEXT DEFAULT 'user',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Tabela de alocação de salas
CREATE TABLE IF NOT EXISTS public.alocacao (
  id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  num_seq TEXT, unidade TEXT, curso TEXT, periodo TEXT,
  turma TEXT, grupo TEXT, codigo TEXT, disciplina TEXT,
  professor TEXT, dia_semana TEXT, horario TEXT,
  sala TEXT, capacidade TEXT, raw JSONB,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. Tabela de eventos
CREATE TABLE IF NOT EXISTS public.eventos (
  id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  data TEXT, dia_semana TEXT, nome TEXT, responsavel TEXT,
  sala TEXT, capacidade TEXT, horario_inicio TEXT,
  horario_termino TEXT, observacoes TEXT, assistente TEXT,
  criado_em TIMESTAMPTZ DEFAULT NOW()
);

-- 4. Tabela de cancelamentos
CREATE TABLE IF NOT EXISTS public.cancelamentos (
  id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
  sala TEXT, dia TEXT, horario TEXT,
  disciplina TEXT, tipo TEXT DEFAULT 'semana',
  motivo TEXT, criado_por UUID REFERENCES auth.users,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 5. Habilitar RLS
ALTER TABLE public.user_profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.alocacao ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.eventos ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.cancelamentos ENABLE ROW LEVEL SECURITY;

-- Políticas
CREATE POLICY "read_alocacao" ON public.alocacao FOR SELECT USING (true);
CREATE POLICY "manage_alocacao" ON public.alocacao FOR ALL USING (auth.role() = 'authenticated');
CREATE POLICY "read_eventos" ON public.eventos FOR SELECT USING (true);
CREATE POLICY "manage_eventos" ON public.eventos FOR ALL USING (auth.role() = 'authenticated');
CREATE POLICY "read_cancel" ON public.cancelamentos FOR SELECT USING (true);
CREATE POLICY "manage_cancel" ON public.cancelamentos FOR ALL USING (auth.role() = 'authenticated');
CREATE POLICY "own_profile" ON public.user_profiles FOR SELECT USING (auth.uid() = id);
CREATE POLICY "admin_all_profiles" ON public.user_profiles FOR ALL USING (true);

-- Trigger para criar perfil automaticamente
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
  INSERT INTO public.user_profiles (id, email) VALUES (NEW.id, NEW.email) ON CONFLICT (id) DO NOTHING;
  RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
  AFTER INSERT ON auth.users
  FOR EACH ROW EXECUTE PROCEDURE public.handle_new_user();`

  function copySQL() { navigator.clipboard.writeText(SQL).catch(() => {}); alert('SQL copiado!') }

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20 }}>
        <h2 style={{ fontSize: 22, color: 'var(--navy)' }}>Relatórios</h2>
        <button className="btn btn-ghost" onClick={() => setSqlModal(true)}><Icon.Code />Setup do Banco</button>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(175px,1fr))', gap: 14, marginBottom: 24 }}>
        {[
          { num: salasUnicas.length, label: 'Salas Cadastradas' },
          { num: profs.length, label: 'Professores' },
          { num: cursos.length, label: 'Cursos' },
          { num: aulas.length, label: 'Total de Alocações' },
          { num: eventos.length, label: 'Eventos' },
          { num: cancelamentos.length, label: 'Cancelamentos' },
        ].map((s, i) => (
          <div key={i} className="stat-box"><div className="stat-num">{s.num}</div><div className="stat-label">{s.label}</div></div>
        ))}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(300px,1fr))', gap: 20 }}>
        <div className="card">
          <h3 style={{ fontSize: 15, marginBottom: 16, color: 'var(--navy)' }}>Aulas por Dia da Semana</h3>
          {aulasPerDia.map((d, i) => (
            <div key={i} style={{ marginBottom: 10 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                <span style={{ fontSize: 13 }}>{d.dia}</span>
                <span style={{ fontSize: 13, fontWeight: 600, color: 'var(--navy)' }}>{d.count}</span>
              </div>
              <div style={{ background: 'var(--gray-100)', borderRadius: 4, height: 8 }}>
                <div style={{ background: 'var(--navy)', width: `${Math.round(d.count / maxD * 100)}%`, height: '100%', borderRadius: 4, transition: 'width .3s' }} />
              </div>
            </div>
          ))}
        </div>
        <div className="card">
          <h3 style={{ fontSize: 15, marginBottom: 16, color: 'var(--navy)' }}>Top 5 Salas Mais Ocupadas</h3>
          {top5Salas.map((s, i) => (
            <div key={i} style={{ marginBottom: 10 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                <span style={{ fontSize: 13, fontWeight: 600 }}>{s.sala}</span>
                <span style={{ fontSize: 13, color: 'var(--gray-600)' }}>{s.count} aulas/sem</span>
              </div>
              <div style={{ background: 'var(--gray-100)', borderRadius: 4, height: 8 }}>
                <div style={{ background: 'var(--gold)', width: `${Math.round(s.count / maxS * 100)}%`, height: '100%', borderRadius: 4 }} />
              </div>
            </div>
          ))}
          {top5Salas.length === 0 && <p style={{ color: 'var(--gray-400)', fontSize: 14 }}>Sem dados</p>}
        </div>
        {cursosDist.length > 0 && (
          <div className="card">
            <h3 style={{ fontSize: 15, marginBottom: 16, color: 'var(--navy)' }}>Distribuição por Curso</h3>
            {cursosDist.map((c, i) => {
              const max = Math.max(...cursosDist.map(x => x.count), 1)
              return (
                <div key={i} style={{ marginBottom: 10 }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                    <span style={{ fontSize: 13, fontWeight: 600 }}>{c.abrev}</span>
                    <span style={{ fontSize: 13, color: 'var(--gray-600)' }}>{c.count}</span>
                  </div>
                  <div style={{ background: 'var(--gray-100)', borderRadius: 4, height: 8 }}>
                    <div style={{ background: 'var(--info)', width: `${Math.round(c.count / max * 100)}%`, height: '100%', borderRadius: 4 }} />
                  </div>
                </div>
              )
            })}
          </div>
        )}
      </div>
      {sqlModal && (
        <div className="modal-overlay">
          <div className="modal modal-lg">
            <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 16 }}>
              <h3 style={{ color: 'var(--navy)' }}>Setup do Banco de Dados (Supabase)</h3>
              <button className="btn btn-ghost btn-sm" onClick={() => setSqlModal(false)}><Icon.X /></button>
            </div>
            <p style={{ fontSize: 13, color: 'var(--gray-600)', marginBottom: 12 }}>Execute este SQL no <strong>SQL Editor</strong> do seu projeto Supabase para criar as tabelas e políticas necessárias.</p>
            <div className="code-block" style={{ fontSize: 11, maxHeight: 400, overflowY: 'auto' }}>{SQL}</div>
            <button className="btn btn-primary" style={{ marginTop: 14 }} onClick={copySQL}><Icon.Copy />Copiar SQL</button>
          </div>
        </div>
      )}
    </div>
  )
}

/* ── Modal Evento ─────────────────────────────────────────────── */
function ModalEvento({ onClose, onSave, toast }) {
  const [form, setForm] = useState({ nome: '', data: '', dia_semana: '', sala: '', horarioInicio: '', horarioTermino: '', responsavel: '', observacoes: '' })
  const [saving, setSaving] = useState(false)
  async function save() {
    if (!form.nome || !form.sala) { toast('Nome e sala são obrigatórios', 'error'); return }
    setSaving(true)
    const { error } = await db.from('eventos').insert([{
      nome: form.nome, data: form.data, dia_semana: form.dia_semana, sala: form.sala,
      horario_inicio: form.horarioInicio, horario_termino: form.horarioTermino,
      responsavel: form.responsavel, observacoes: form.observacoes,
    }])
    if (error) toast(error.message, 'error')
    else { toast('Evento publicado!', 'success'); onSave(); onClose() }
    setSaving(false)
  }
  return (
    <div className="modal-overlay">
      <div className="modal">
        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 20 }}>
          <h3 style={{ color: 'var(--navy)' }}>Publicar Evento</h3>
          <button className="btn btn-ghost btn-sm" onClick={onClose}><Icon.X /></button>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          {[['Nome do Evento *', 'nome', 'text'], ['Data', 'data', 'date'], ['Sala *', 'sala', 'text'],
            ['Horário Início', 'horarioInicio', 'time'], ['Horário Término', 'horarioTermino', 'time'], ['Responsável', 'responsavel', 'text'],
          ].map(([lbl, key, type]) => (
            <div key={key}><label className="label">{lbl}</label>
              <input className="input" type={type} value={form[key]} onChange={e => setForm({ ...form, [key]: e.target.value })} />
            </div>
          ))}
          <div><label className="label">Dia da Semana</label>
            <select className="input" value={form.dia_semana} onChange={e => setForm({ ...form, dia_semana: e.target.value })}>
              <option value="">Selecione...</option>
              {DIAS.map(d => <option key={d} value={d}>{DIAS_FULL[d]}</option>)}
            </select>
          </div>
          <div><label className="label">Observações</label>
            <textarea className="input" value={form.observacoes} onChange={e => setForm({ ...form, observacoes: e.target.value })} rows={3} style={{ resize: 'vertical' }} />
          </div>
          <button className="btn btn-gold" onClick={save} disabled={saving}>
            {saving ? <><span className="loading-spin">↻</span> Publicando...</> : 'Publicar Evento'}
          </button>
        </div>
      </div>
    </div>
  )
}

/* ── Configurações ────────────────────────────────────────────── */
function Configuracoes({ currentUser, toast }) {
  const [senha, setSenha] = useState({ nova: '', confirma: '' })
  const [loading, setLoading] = useState(false)

  async function changePass() {
    if (senha.nova !== senha.confirma) { toast('As senhas não conferem', 'error'); return }
    if (senha.nova.length < 6) { toast('Senha deve ter no mínimo 6 caracteres', 'error'); return }
    setLoading(true)
    const { error } = await db.auth.updateUser({ password: senha.nova })
    if (error) toast(error.message, 'error')
    else { toast('Senha alterada com sucesso!', 'success'); setSenha({ nova: '', confirma: '' }) }
    setLoading(false)
  }

  return (
    <div>
      <h2 style={{ marginBottom: 20, fontSize: 22, color: 'var(--navy)' }}>Configurações</h2>
      <div style={{ maxWidth: 480, display: 'flex', flexDirection: 'column', gap: 16 }}>
        <div className="card">
          <h3 style={{ fontSize: 16, marginBottom: 16, color: 'var(--navy)' }}>Alterar Senha</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
            <div><label className="label">Nova senha</label>
              <input className="input" type="password" value={senha.nova} onChange={e => setSenha({ ...senha, nova: e.target.value })} placeholder="••••••••" />
            </div>
            <div><label className="label">Confirmar nova senha</label>
              <input className="input" type="password" value={senha.confirma} onChange={e => setSenha({ ...senha, confirma: e.target.value })} placeholder="••••••••" />
            </div>
            <button className="btn btn-primary" onClick={changePass} disabled={loading}>
              {loading ? <><span className="loading-spin">↻</span> Salvando...</> : 'Salvar Senha'}
            </button>
          </div>
        </div>
        <div className="card">
          <h3 style={{ fontSize: 16, marginBottom: 8, color: 'var(--navy)' }}>Informações da Conta</h3>
          <p style={{ fontSize: 14, color: 'var(--gray-600)' }}>E-mail: <strong>{currentUser?.email}</strong></p>
          <p style={{ fontSize: 12, color: 'var(--gray-400)', marginTop: 4 }}>ID: {currentUser?.id}</p>
        </div>
        <div className="card" style={{ border: '1px solid var(--gray-200)' }}>
          <h3 style={{ fontSize: 15, marginBottom: 8, color: 'var(--navy)' }}>Modo Totem</h3>
          <p style={{ fontSize: 13, color: 'var(--gray-600)', marginBottom: 12 }}>Abre a visualização pública do campus em uma nova aba — ideal para TVs e totens.</p>
          <button className="btn btn-ghost" onClick={() => window.open('?modo=totem', '_blank')}><Icon.Monitor />Abrir Totem</button>
        </div>
      </div>
    </div>
  )
}

/* ── Totem (Modo Público) ─────────────────────────────────────── */
function TotemView() {
  const [aulas, setAulas] = useState([])
  const [eventos, setEventos] = useState([])
  const [cancelamentos, setCancelamentos] = useState([])
  const [clock, setClock] = useState(nowTimeWithSec())
  const [loading, setLoading] = useState(true)
  const hoje = todayDia()

  async function loadData() {
    const [a, e, c] = await Promise.all([
      db.from('alocacao').select('*'),
      db.from('eventos').select('*').order('data', { ascending: true }),
      db.from('cancelamentos').select('*'),
    ])
    const rows = (a.data || []).map(d => d.raw || {
      'DIA DA SEMANA': d.dia_semana, 'HORÁRIO': d.horario, 'SALA': d.sala,
      'DISCIPLINA': d.disciplina, 'PROFESSOR': d.professor, 'CURSO': d.curso,
      'PERÍODO': d.periodo, 'TURMA': d.turma, 'GRUPO': d.grupo, 'CAPACIDADE': d.capacidade,
      ...d.raw || {}
    })
    setAulas(rows)
    setEventos((e.data || []).map(d => ({ 'Nome do Evento': d.nome, 'Data': d.data, 'Dia da Semana': d.dia_semana, 'Sala': d.sala, 'Horário Inicio': d.horario_inicio, 'Horario Termino': d.horario_termino, 'Responsável Reserva': d.responsavel })))
    setCancelamentos(c.data || [])
    setLoading(false)
  }

  useEffect(() => {
    loadData()
    const tick = setInterval(() => setClock(nowTimeWithSec()), 1000)
    const refresh = setInterval(loadData, 300000)
    return () => { clearInterval(tick); clearInterval(refresh) }
  }, [])

  const isCancelado = row => cancelamentos.some(c =>
    normalize(c.sala) === normalize(row['SALA']) &&
    normalize(c.horario) === normalize(row['HORÁRIO']) &&
    normalize(c.dia) === normalize(row['DIA DA SEMANA'])
  )

  const aulasDia = useMemo(() => {
    const map = new Map()
    aulas.filter(a => normalize(a['DIA DA SEMANA']) === normalize(hoje)).forEach(row => {
      const key = `${normalize(row['SALA'])}|${normalize(row['HORÁRIO'])}|${normalize(row['DISCIPLINA'])}`
      if (!map.has(key)) map.set(key, { ...row, _codigos: [buildCodigo(row)], _cancelado: isCancelado(row) })
      else { const ex = map.get(key); const c = buildCodigo(row); if (!ex._codigos.includes(c)) ex._codigos.push(c) }
    })
    return [...map.values()]
  }, [aulas, hoje, cancelamentos])

  const bySala = useMemo(() => {
    const m = new Map()
    aulasDia.forEach(a => { const s = a['SALA'] || 'Sem sala'; if (!m.has(s)) m.set(s, []); m.get(s).push(a) })
    return m
  }, [aulasDia])

  const eventosDia = useMemo(() => eventos.filter(e => {
    const dds = normalize(e['Dia da Semana'] || '')
    return dds === normalize(hoje) || dds === normalize(DIAS_LABEL[hoje] || '')
  }), [eventos, hoje])

  const agora = new Date()
  const dateStr = agora.toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' })

  if (loading) return (
    <div className="totem-body" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
      <div style={{ textAlign: 'center' }}>
        <div style={{ fontFamily: 'Sora', fontSize: 48, fontWeight: 800, color: 'var(--gold)', marginBottom: 16 }}>ibmec</div>
        <div className="loading-spin" style={{ fontSize: 28, color: 'rgba(255,255,255,.5)' }}>↻</div>
      </div>
    </div>
  )

  return (
    <div className="totem-body">
      <div className="totem-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          <div style={{ width: 12, height: 12, borderRadius: '50%', background: 'var(--gold)' }} />
          <span style={{ fontFamily: 'Sora', fontWeight: 800, fontSize: 32, color: '#fff' }}>ibmec</span>
          <div style={{ width: 1, height: 32, background: 'rgba(255,255,255,.2)' }} />
          <span style={{ fontSize: 16, color: 'rgba(255,255,255,.7)', textTransform: 'capitalize' }}>{dateStr}</span>
        </div>
        <div style={{ fontFamily: 'Sora', fontSize: 36, fontWeight: 700, color: 'var(--gold)', letterSpacing: 4 }}>{clock}</div>
      </div>
      <div style={{ padding: '28px 40px', maxWidth: 1600, margin: '0 auto' }}>
        {eventosDia.length > 0 && (
          <div style={{ marginBottom: 24 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--gold)', letterSpacing: 2, textTransform: 'uppercase', marginBottom: 12 }}>Eventos de Hoje</div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12 }}>
              {eventosDia.map((ev, i) => (
                <div key={i} style={{ background: 'rgba(245,172,0,.12)', border: '1.5px solid var(--gold)', borderRadius: 12, padding: '12px 18px', minWidth: 220 }}>
                  <div style={{ fontFamily: 'Sora', fontWeight: 700, fontSize: 16, color: '#fff', marginBottom: 4 }}>{ev['Nome do Evento'] || 'Evento'}</div>
                  <div style={{ fontSize: 13, color: 'rgba(255,255,255,.6)' }}>{ev['Sala'] || ''} · {ev['Horário Inicio'] || ''}{ev['Horario Termino'] ? ' – ' + ev['Horario Termino'] : ''}</div>
                  {ev['Responsável Reserva'] && <div style={{ fontSize: 12, color: 'rgba(255,255,255,.4)', marginTop: 2 }}>{ev['Responsável Reserva']}</div>}
                </div>
              ))}
            </div>
          </div>
        )}
        <div style={{ fontSize: 13, fontWeight: 700, color: 'var(--gold)', letterSpacing: 2, textTransform: 'uppercase', marginBottom: 16 }}>Salas Hoje — {DIAS_FULL[hoje]}</div>
        {bySala.size === 0 ? (
          <div style={{ textAlign: 'center', padding: 80, color: 'rgba(255,255,255,.3)', fontSize: 18 }}>Nenhuma aula cadastrada para hoje</div>
        ) : (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill,minmax(320px,1fr))', gap: 16 }}>
            {[...bySala.entries()].sort((a, b) => a[0] > b[0] ? 1 : -1).map(([sala, list]) => (
              <div key={sala} style={{ background: 'rgba(255,255,255,.04)', border: '1px solid rgba(255,255,255,.08)', borderRadius: 14, overflow: 'hidden' }}>
                <div style={{ background: 'var(--navy-light)', padding: '12px 18px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid var(--gold)' }}>
                  <span style={{ fontFamily: 'Sora', fontWeight: 700, fontSize: 18, color: '#fff' }}>{sala}</span>
                  <span style={{ fontSize: 12, color: 'rgba(255,255,255,.5)' }}>{list.length} aula{list.length !== 1 ? 's' : ''}</span>
                </div>
                <div style={{ padding: '12px 16px' }}>
                  {list.sort((a, b) => (a['HORÁRIO'] || '') > (b['HORÁRIO'] || '') ? 1 : -1).map((a, i) => {
                    const live = isAoVivo(a['HORÁRIO'])
                    return (
                      <div key={i} style={{ background: a._cancelado ? 'rgba(220,38,38,.1)' : live ? 'rgba(22,163,74,.1)' : 'rgba(255,255,255,.04)', borderLeft: `3px solid ${a._cancelado ? 'var(--danger)' : live ? 'var(--success)' : 'var(--gold)'}`, borderRadius: 8, padding: '10px 14px', marginBottom: 10, opacity: a._cancelado ? .45 : 1 }}>
                        <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 4 }}>
                          <span style={{ fontWeight: 700, fontSize: 13, color: 'rgba(255,255,255,.8)' }}>{a['HORÁRIO']}</span>
                          {live && !a._cancelado && <span style={{ fontSize: 11, color: 'var(--success)', fontWeight: 700 }}>● AO VIVO</span>}
                          {a._cancelado && <span style={{ fontSize: 11, color: 'var(--danger)', fontWeight: 700 }}>✕ CANCELADO</span>}
                        </div>
                        <div style={{ fontFamily: 'Sora', fontWeight: 600, fontSize: 15, color: '#fff', marginBottom: 2 }}>{a['DISCIPLINA']}</div>
                        <div style={{ fontSize: 12, color: 'rgba(255,255,255,.5)', marginBottom: 6 }}>{a['PROFESSOR']}</div>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4 }}>
                          {a._codigos.map((c, ci) => <span key={ci} style={{ background: 'var(--navy)', color: 'var(--gold)', padding: '2px 8px', borderRadius: 5, fontSize: 11, fontWeight: 700 }}>{c}</span>)}
                        </div>
                      </div>
                    )
                  })}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
      <div style={{ textAlign: 'center', padding: '16px 0', fontSize: 11, color: 'rgba(255,255,255,.2)' }}>
        Atualizado automaticamente a cada 5 minutos · Ibmec Campus System
      </div>
    </div>
  )
}

/* ── Main App ─────────────────────────────────────────────────── */
export default function App() {
  const params = new URLSearchParams(window.location.search)
  const isTotem = params.get('modo') === 'totem'
  if (isTotem) return <TotemView />

  const [user, setUser] = useState(null)
  const [userRole, setUserRole] = useState('user')
  const [aulas, setAulas] = useState([])
  const [eventos, setEventos] = useState([])
  const [cancelamentos, setCancelamentos] = useState([])
  const [page, setPage] = useState('dashboard')
  const [toast, setToast] = useState(null)
  const [eventoModal, setEventoModal] = useState(false)
  const [sidebarCollapsed, setSidebarCollapsed] = useState(false)
  const [loading, setLoading] = useState(true)

  function showToast(msg, type = 'success') { setToast({ msg, type }) }

  useEffect(() => {
    db.auth.getSession().then(({ data: { session } }) => {
      if (session?.user) { setUser(session.user); loadProfile(session.user.id) }
      setLoading(false)
    })
    const { data: { subscription } } = db.auth.onAuthStateChange((_, session) => {
      setUser(session?.user || null)
      if (session?.user) loadProfile(session.user.id)
    })
    return () => subscription.unsubscribe()
  }, [])

  async function loadProfile(uid) {
    const { data } = await db.from('user_profiles').select('role').eq('id', uid).single()
    setUserRole(data?.role || 'user')
  }

  useEffect(() => {
    if (!user) return
    loadAulas(); loadEventos(); loadCancelamentos()
    const t = setInterval(loadEventos, 3600000)
    return () => clearInterval(t)
  }, [user])

  async function loadAulas() {
    const { data } = await db.from('alocacao').select('*')
    if (data) setAulas(data.map(d => d.raw || {
      'DIA DA SEMANA': d.dia_semana, 'HORÁRIO': d.horario, 'SALA': d.sala,
      'DISCIPLINA': d.disciplina, 'PROFESSOR': d.professor, 'CURSO': d.curso,
      'PERÍODO': d.periodo, 'TURMA': d.turma, 'GRUPO': d.grupo, 'CAPACIDADE': d.capacidade,
      ...(d.raw || {})
    }))
  }
  async function loadEventos() {
    const { data } = await db.from('eventos').select('*').order('data', { ascending: true })
    if (data) setEventos(data.map(d => ({
      'Nome do Evento': d.nome, 'Data': d.data, 'Dia da Semana': d.dia_semana,
      'Sala': d.sala, 'Horário Inicio': d.horario_inicio, 'Horario Termino': d.horario_termino,
      'Responsável Reserva': d.responsavel, 'Observações': d.observacoes,
      nome: d.nome, data: d.data, dia_semana: d.dia_semana, sala: d.sala,
      horario_inicio: d.horario_inicio, horario_termino: d.horario_termino, responsavel: d.responsavel,
    })))
  }
  async function loadCancelamentos() {
    const { data } = await db.from('cancelamentos').select('*')
    if (data) setCancelamentos(data)
  }

  async function handleLogout() {
    await db.auth.signOut()
    setUser(null); setAulas([]); setEventos([]); setCancelamentos([])
  }

  if (loading) return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'var(--navy)' }}>
      <div style={{ textAlign: 'center', color: '#fff' }}>
        <div style={{ fontFamily: 'Sora', fontSize: 36, fontWeight: 800, marginBottom: 12, color: 'var(--gold)' }}>ibmec</div>
        <div className="loading-spin" style={{ fontSize: 26 }}>↻</div>
      </div>
    </div>
  )

  if (!user) return <LoginPage onLogin={u => { setUser(u); loadProfile(u.id) }} />

  const isAdmin = userRole === 'admin' || userRole === 'sub-admin'
  const sideW = sidebarCollapsed ? 64 : 220

  function renderPage() {
    switch (page) {
      case 'dashboard':     return <Dashboard aulas={aulas} eventos={eventos} cancelamentos={cancelamentos} />
      case 'salas':         return <SalasHoje aulas={aulas} eventos={eventos} cancelamentos={cancelamentos} />
      case 'grade':         return <GradeSemanal aulas={aulas} cancelamentos={cancelamentos} />
      case 'eventos':       return <Eventos eventos={eventos} isAdmin={isAdmin} onAddEvento={() => setEventoModal(true)} />
      case 'planilhas':     return isAdmin ? <Planilhas onAlocacaoUpload={rows => setAulas(rows)} onEventosUpload={rows => setEventos(rows)} toast={showToast} /> : null
      case 'cancelamentos': return isAdmin ? <CancelamentosAdmin toast={showToast} /> : null
      case 'usuarios':      return isAdmin ? <Usuarios toast={showToast} currentUser={user} /> : null
      case 'relatorios':    return isAdmin ? <Relatorios aulas={aulas} eventos={eventos} cancelamentos={cancelamentos} /> : null
      case 'configuracoes': return <Configuracoes currentUser={user} toast={showToast} />
      default:              return <Dashboard aulas={aulas} eventos={eventos} cancelamentos={cancelamentos} />
    }
  }

  return (
    <div style={{ display: 'flex', minHeight: '100vh' }}>
      <Sidebar page={page} setPage={p => { setPage(p); if (p === 'cancelamentos') loadCancelamentos() }} isAdmin={isAdmin} onLogout={handleLogout} collapsed={sidebarCollapsed} />
      <div style={{ flex: 1, marginLeft: sideW, transition: 'margin .2s', display: 'flex', flexDirection: 'column', minHeight: '100vh' }}>
        <div style={{ height: 56, background: '#fff', borderBottom: '1px solid var(--gray-200)', display: 'flex', alignItems: 'center', padding: '0 24px', gap: 12, position: 'sticky', top: 0, zIndex: 50 }}>
          <button className="btn btn-ghost btn-sm" onClick={() => setSidebarCollapsed(!sidebarCollapsed)}>
            <span style={{ width: 18, height: 18, display: 'flex' }}><Icon.Menu /></span>
          </button>
          <div style={{ flex: 1 }} />
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <span className={`badge badge-${userRole === 'admin' ? 'navy' : userRole === 'sub-admin' ? 'gold' : 'gray'}`}>{userRole}</span>
            <span style={{ fontSize: 13, color: 'var(--gray-600)' }}>{user.email}</span>
          </div>
        </div>
        <div style={{ flex: 1, padding: 28, overflowY: 'auto' }}>
          {renderPage()}
        </div>
      </div>
      {eventoModal && <ModalEvento onClose={() => setEventoModal(false)} onSave={loadEventos} toast={showToast} />}
      {toast && <Toast msg={toast.msg} type={toast.type} onClose={() => setToast(null)} />}
    </div>
  )
}
