export const CURSO_MAP = {
  'ADMINISTRAÇÃO': 'ADM',
  'CIÊNCIA DE DADOS E INTELIGÊNCIA ARTIFICIAL': 'CN E IA',
  'CIÊNCIAS ECONÔMICAS': 'ECO',
  'ENGENHARIA DE COMPUTAÇÃO': 'ENG COMP',
  'ENGENHARIA DE PRODUÇÃO': 'ENG PROD',
  'ENGENHARIA DE SOFTWARE': 'ENG SOFT',
}

export function abrevCurso(c) {
  if (!c) return ''
  const u = c.toUpperCase().trim()
  return CURSO_MAP[u] || c
}

export function buildCodigo(row) {
  const per = (row['PERÍODO'] || '').toString().replace(/[ºo°]/g, '').trim() + 'º'
  const cur = abrevCurso(row['CURSO'] || '')
  const tur = (row['TURMA'] || '').toString().trim()
  const grp = (row['GRUPO'] || '').toString().trim()
  return `${per} ${cur} ${tur} ${grp}`.trim()
}

export function normalize(s) {
  return (s || '').toString().trim().toUpperCase()
}

export const DIAS = ['SEGUNDA', 'TERÇA', 'QUARTA', 'QUINTA', 'SEXTA', 'SÁBADO', 'DOMINGO']
export const DIAS_FULL = {
  'SEGUNDA': 'Segunda-feira', 'TERÇA': 'Terça-feira', 'QUARTA': 'Quarta-feira',
  'QUINTA': 'Quinta-feira', 'SEXTA': 'Sexta-feira', 'SÁBADO': 'Sábado', 'DOMINGO': 'Domingo',
}
export const DIAS_LABEL = {
  'SEGUNDA': 'Seg', 'TERÇA': 'Ter', 'QUARTA': 'Qua',
  'QUINTA': 'Qui', 'SEXTA': 'Sex', 'SÁBADO': 'Sáb', 'DOMINGO': 'Dom',
}

export function todayDia() {
  const m = { 0: 'DOMINGO', 1: 'SEGUNDA', 2: 'TERÇA', 3: 'QUARTA', 4: 'QUINTA', 5: 'SEXTA', 6: 'SÁBADO' }
  return m[new Date().getDay()] || 'SEGUNDA'
}

export function fmtDate(d) {
  if (!d) return '-'
  try {
    const dt = new Date(d)
    if (isNaN(dt)) return d
    return dt.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit', year: 'numeric' })
  } catch { return d }
}

export function nowTime() {
  const n = new Date()
  return `${String(n.getHours()).padStart(2, '0')}:${String(n.getMinutes()).padStart(2, '0')}`
}

export function nowTimeWithSec() {
  const n = new Date()
  return `${String(n.getHours()).padStart(2, '0')}:${String(n.getMinutes()).padStart(2, '0')}:${String(n.getSeconds()).padStart(2, '0')}`
}

export function isAoVivo(horario) {
  const h = horario || ''
  const parts = h.split('-')
  if (parts.length < 2) return false
  const t = nowTime()
  return parts[0].trim() <= t && t <= parts[1].trim()
}
