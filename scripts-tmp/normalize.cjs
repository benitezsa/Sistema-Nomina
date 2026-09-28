'use strict'

const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']

function round2(n) {
  return Math.round(Number(n) * 100) / 100
}

function fechaASerial(v) {
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v.toISOString().slice(0, 10)
  if (typeof v === 'number') {
    const excelEpoch = Date.UTC(1899, 11, 30)
    return new Date(excelEpoch + v * 86400000).toISOString().slice(0, 10)
  }
  if (typeof v === 'string') {
    const t = v.trim()
    if (!t) return ''
    const d = new Date(t)
    return Number.isNaN(d.getTime()) ? t : d.toISOString().slice(0, 10)
  }
  if (typeof v === 'object' && v !== null) {
    if ('result' in v) return fechaASerial(v.result)
    if ('text' in v) return fechaASerial(v.text)
  }
  return ''
}

function fechaADate(v) {
  if (!v) return null
  if (v instanceof Date) return Number.isNaN(v.getTime()) ? null : v
  if (typeof v === 'number') {
    const excelEpoch = Date.UTC(1899, 11, 30)
    return new Date(excelEpoch + v * 86400000)
  }
  if (typeof v === 'string') {
    const d = new Date(v.trim())
    return Number.isNaN(d.getTime()) ? null : d
  }
  if (typeof v === 'object' && v !== null) {
    if ('result' in v) return fechaADate(v.result)
    if ('text' in v) return fechaADate(v.text)
  }
  return null
}

function esDiaLaborable(fecha, jornada) {
  const dow = fecha.getUTCDay()
  if (jornada === 6) return dow !== 0
  if (jornada === 5) return dow !== 0 && dow !== 6
  return dow >= 1 && dow <= 4
}

function laborablesRango(desde, hasta, jornada, feriados) {
  if (!desde || !hasta || hasta < desde) return 0
  const inicio = new Date(Date.UTC(desde.getUTCFullYear(), desde.getUTCMonth(), desde.getUTCDate()))
  const fin = new Date(Date.UTC(hasta.getUTCFullYear(), hasta.getUTCMonth(), hasta.getUTCDate()))
  let total = 0
  const cur = new Date(inicio)
  while (cur <= fin) {
    const iso = cur.toISOString().slice(0, 10)
    if (esDiaLaborable(cur, jornada) && !feriados.has(iso)) total++
    cur.setUTCDate(cur.getUTCDate() + 1)
  }
  return total
}

function calcularNoLaborados(incidencias, jornada, feriados) {
  return incidencias.reduce((total, inc) => {
    if (inc.diasNeto !== undefined && inc.diasNeto !== null) return total + inc.diasNeto
    if (inc.fechaInicio && inc.fechaFin) {
      return total + laborablesRango(inc.fechaInicio, inc.fechaFin, jornada, feriados)
    }
    return total + (inc.cantidadDias ?? 0)
  }, 0)
}

function derivar(w) {
  const jornada = w.jornada ?? 5
  if (typeof jornada !== 'number') {
    return { noLab: null, posibles: null, ef: null, remT: null }
  }
  const incidencias = (w.incidencias || []).map((i) => ({
    cantidadDias: i.cantidadDias,
    diasNeto: i.diasNeto,
    fechaInicio: fechaADate(i.fechaInicio),
    fechaFin: fechaADate(i.fechaFin),
  }))
  const noLaborados = calcularNoLaborados(incidencias, jornada, new Set(w.feriados || []))
  const mesesPresentes = new Set(w.mesesPresentes || [])
  let posibles
  if (w.diasPosibles != null) {
    posibles = w.diasPosibles
  } else if (w.laborablesPorMes) {
    posibles = months.reduce((s, m, i) => s + (mesesPresentes.has(m) ? w.laborablesPorMes[m] || 0 : 0), 0)
  } else {
    posibles = w.diasTrabajados ? w.diasTrabajados.total || 0 : 0
  }
  const efectivos = Math.max(0, posibles - noLaborados)
  return {
    noLab: round2(noLaborados),
    posibles: round2(posibles),
    ef: round2(efectivos),
    remT: round2(w.remuneraciones ? w.remuneraciones.total || 0 : 0),
  }
}

function normalizar(w) {
  const out = {
    dni: String(w.dni),
    apP: w.apellidoPaterno || '',
    apM: w.apellidoMaterno || '',
    nom: w.nombres || '',
    fi: fechaASerial(w.fechaInicio),
    fc: fechaASerial(w.fechaCese),
  }
  if (w.remuneraciones) {
    out.rem = { m: months.map((m) => round2(w.remuneraciones[m])), t: round2(w.remuneraciones.total) }
  }
  if (w.diasTrabajados) {
    out.dias = { m: months.map((m) => round2(w.diasTrabajados[m])), t: round2(w.diasTrabajados.total) }
  }
  out.mRem = (w.mesesRemuneraciones || []).slice().sort()
  out.mDias = (w.mesesDias || []).slice().sort()
  out.mPres = (w.mesesPresentes || []).slice().sort()
  out.jornada = w.jornada ?? null
  out.pos = w.diasPosibles != null ? round2(w.diasPosibles) : null
  out.noLabRef = w.diasNoLaboradosReferencia != null ? round2(w.diasNoLaboradosReferencia) : null
  out.efRef = w.diasEfectivosReferencia != null ? round2(w.diasEfectivosReferencia) : null
  if (w.laborablesPorMes) {
    out.lab = { m: months.map((m) => round2(w.laborablesPorMes[m])), t: round2(w.laborablesPorMes.total) }
  }
  out.feriados = (w.feriados || []).slice().sort()
  out.inc = (w.incidencias || [])
    .map((i) => ({
      cod: i.codigo || '',
      desc: i.descripcion || '',
      cant: round2(i.cantidadDias),
      neto: i.diasNeto != null ? round2(i.diasNeto) : null,
      fi: fechaASerial(i.fechaInicio),
      ff: fechaASerial(i.fechaFin),
    }))
    .sort((a, b) => a.desc.localeCompare(b.desc) || a.cant - b.cant || a.cod.localeCompare(b.cod))
  out.derivado = derivar(w)
  return out
}

function normalizarLista(workers) {
  return workers.map(normalizar).sort((a, b) => a.dni.localeCompare(b.dni))
}

module.exports = { normalizar, normalizarLista, months, round2, calcularNoLaborados, fechaASerial }