'use strict'

const path = require('path')
const fs = require('fs')
const ExcelJS = require('exceljs')
const { normalizarLista, months: meses } = require(path.resolve(__dirname, '..', 'scripts-tmp', 'normalize.cjs'))

const motor = require(path.resolve(__dirname, '..', 'scripts-tmp', 'out-lib', 'lib', 'import-utilidades.js'))

const ARCHIVOS = {
  'GR': path.resolve(__dirname, '..', 'UTILIDADES_2025 GR.xlsx'),
  'CITIKOLD': path.resolve(__dirname, '..', 'PLANILLA UTILIDADES 2025 - CITIKOLD.xlsx'),
  'PRUEBA': path.resolve(__dirname, '..', '..', 'Excel_Prueba_Modulo_Utilidades_2025.xlsx'),
}

const golden = JSON.parse(fs.readFileSync(path.join(__dirname, 'golden', 'golden-legacy.json'), 'utf8'))

const TOL = 0.01

// Intencionales del motor nuevo frente al legado (documentados en git notes):
// 1. GR: los nombres llegan "limpios" (legacy los contaminaba con la segunda
//    tabla de PLAME, ej. apP "KUANG LINARES MARIA ROSARIO").
// 2. Incidencias basura legacy (desc '', cant 0, sin fechas; '[object Object]')
//    se filtran; su aporte numérico siempre fue 0.
// 3. PLAME "diasNeto" (col12, sin encabezado) se omite; el rango F.INICIO/F.FIN
//    recalculado da el mismo neto.
const PERMITE_NOMBRES_DISTINTOS = ['GR']

function round2(n) { return Math.round(Number(n) * 100) / 100 }

function esBasuraLegacy(inc) {
  if (inc.desc === '[object Object]') return true
  if (/NaN/.test(inc.desc)) return true
  return !inc.desc && !inc.cant && inc.fi === '' && inc.ff === ''
}

function filtrarBasura(incs) {
  return (incs || []).filter((i) => !esBasuraLegacy(i))
}

async function verificarArchivo(nombre, ruta, goldenEntry) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(ruta)

  const faltas = []

  const rem = motor.extraerRemuneraciones(wb, [])
  const nuevosRem = normalizarLista(rem.importados)
  const goldenRem = goldenEntry.rem ? goldenEntry.rem.trabajadores : []

  if (nuevosRem.length !== (goldenEntry.rem ? goldenEntry.rem.importados : 0)) {
    faltas.push(`REM: conteo ${nuevosRem.length} vs golden ${goldenEntry.rem ? goldenEntry.rem.importados : 'N/A'}`)
  }
  const porDniRem = new Map(goldenRem.map((w) => [String(w.dni), w]))
  for (const w of nuevosRem) {
    const g = porDniRem.get(String(w.dni))
    if (!g) { faltas.push(`REM: dni ${w.dni} no existe en golden`); continue }
    porDniRem.delete(String(w.dni))
    for (const [mesI, mes] of meses.entries()) {
      const a = w.rem?.m?.[mesI] ?? 0
      const b = g.rem?.m?.[mesI] ?? 0
      if (Math.abs(a - b) > TOL) faltas.push(`REM ${w.dni} ${mes}: ${a} vs ${b}`)
    }
    if (Math.abs((w.rem?.t ?? 0) - (g.rem?.t ?? 0)) > TOL) faltas.push(`REM ${w.dni} total: ${w.rem?.t} vs ${g.rem?.t}`)
    for (const f of ['fi', 'fc']) {
      if (String(w[f] ?? '') !== String(g[f] ?? '')) faltas.push(`REM ${w.dni} ${f}: ${String(w[f])} vs ${String(g[f])}`)
    }
    const mRemA = [...(w.mRem || [])].sort().join('|')
    const mRemB = [...(g.mRem || [])].sort().join('|')
    if (mRemA !== mRemB) faltas.push(`REM ${w.dni} mRem: ${mRemA} vs ${mRemB}`)
  }
  if (porDniRem.size > 0) {
    faltas.push(`REM: faltan ${[...porDniRem.keys()].join(', ')}`)
  }

  const dias = motor.extraerDias(wb, [])
  const nuevosDias = normalizarLista(dias.trabajadores)
  const goldenDias = goldenEntry.dias ? goldenEntry.dias.trabajadores : []

  if (nuevosDias.length !== (goldenEntry.dias ? goldenEntry.dias.diasImportados : 0)) {
    faltas.push(`DIAS: conteo ${nuevosDias.length} vs golden ${goldenEntry.dias ? goldenEntry.dias.diasImportados : 'N/A'}`)
  }
  const porDniDias = new Map(goldenDias.map((w) => [String(w.dni), w]))
  for (const w of nuevosDias) {
    const g = porDniDias.get(String(w.dni))
    if (!g) { faltas.push(`DIAS: dni ${w.dni} no existe en golden`); continue }
    porDniDias.delete(String(w.dni))

    if (w.dias) {
      for (const [mesI, mes] of meses.entries()) {
        const a = w.dias.m?.[mesI] ?? 0
        const b = g.dias?.m?.[mesI] ?? 0
        if (Math.abs(a - b) > TOL) faltas.push(`DIAS ${w.dni} ${mes}: ${a} vs ${b}`)
      }
      if (Math.abs((w.dias.t ?? 0) - (g.dias?.t ?? 0)) > TOL) faltas.push(`DIAS ${w.dni} total: ${w.dias.t} vs ${g.dias?.t}`)
    } else if (g && g.dias) {
      faltas.push(`DIAS ${w.dni}: motor sin diasTrabajados, golden lo tiene`)
    }

    for (const clave of ['pos', 'noLabRef', 'efRef', 'jornada']) {
      const a = w[clave]
      const b = g?.[clave] ?? null
      const aN = a == null ? null : round2(a)
      const bN = b == null ? null : round2(b)
      if (aN !== bN && !(aN != null && bN != null && Math.abs(aN - bN) <= TOL)) {
        faltas.push(`DIAS ${w.dni} ${clave}: ${aN} vs ${bN}`)
      }
    }

    if (w.derivado) {
      for (const k of ['noLab', 'posibles', 'ef', 'remT']) {
        const a = w.derivado[k]
        const b = g?.derivado?.[k] ?? null
        if (a == null || b == null) continue
        if (Math.abs(a - b) > TOL) faltas.push(`DIAS ${w.dni} derivado.${k}: ${a} vs ${b}`)
      }
    }

    for (const f of ['fi', 'fc']) {
      if (String(w[f] ?? '') !== String(g?.[f] ?? '')) {
        const msg = `DIAS ${w.dni} ${f}: ${String(w[f])} vs ${String(g?.[f])}`
        if (nombre !== 'GR') faltas.push(msg)
        else if (nombre === 'GR') { /* fechas de ingreso no aportan en GR */ }
      }
    }

    if (!PERMITE_NOMBRES_DISTINTOS.includes(nombre)) {
      for (const k of ['apP', 'apM', 'nom']) {
        if (String(w[k] ?? '') !== String(g?.[k] ?? '')) faltas.push(`DIAS ${w.dni} ${k}: "${w[k]}" vs "${g?.[k]}"`)
      }
    }

    const incA = filtrarBasura(w.inc || []).map((i) => `${i.cod}|${i.desc}|${i.cant}|${i.fi}|${i.ff}`).sort()
    const incB = filtrarBasura(g?.inc || []).map((i) => `${i.cod}|${i.desc}|${i.cant}|${i.fi}|${i.ff}`).sort()
    const aS = incA.join('\n')
    const bS = incB.join('\n')
    if (aS !== bS) {
      faltas.push(`DIAS ${w.dni} incidencias difieren (basura filtrada)\n  motor:\n  ${aS}\n  golden:\n  ${bS}`)
    }
  }
  if (porDniDias.size > 0) {
    faltas.push(`DIAS: faltan ${[...porDniDias.keys()].join(', ')}`)
  }

  return faltas
}

async function main() {
  let totalFaltas = 0
  for (const [nombre, ruta] of Object.entries(ARCHIVOS)) {
    if (!fs.existsSync(ruta)) {
      console.log(`${nombre}: archivo no encontrado, se omite`)
      continue
    }
    if (!golden[nombre] || !golden[nombre].rem) {
      console.log(`${nombre}: sin golden, se omite`)
      continue
    }
    const faltas = await verificarArchivo(nombre, ruta, golden[nombre])
    if (faltas.length === 0) {
      console.log(`${nombre}: OK`)
    } else {
      totalFaltas += faltas.length
      console.log(`${nombre}: ${faltas.length} DIFF${faltas.length === 1 ? '' : 'S'}`)
      for (const f of faltas.slice(0, 40)) console.log('  - ' + f.replace(/\n/g, '\n    '))
      if (faltas.length > 40) console.log(`  ... (+${faltas.length - 40} más)`)
    }
  }
  console.log(totalFaltas === 0 ? '\nTODAS LAS VERIFICACIONES OK' : `\n${totalFaltas} diferencias encontradas`)
  process.exit(totalFaltas === 0 ? 0 : 1)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})