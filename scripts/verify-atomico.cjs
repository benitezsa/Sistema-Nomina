'use strict'

const path = require('path')
const ExcelJS = require('exceljs')
const motor = require(path.resolve(__dirname, '..', 'scripts-tmp', 'out-lib', 'lib', 'import-utilidades.js'))
const atomico = require(path.resolve(__dirname, '..', 'scripts-tmp', 'out-lib', 'lib', 'import-utilidades-atomico.js'))

let fallos = 0

function check(nombre, condicion, detalle) {
  if (condicion) {
    console.log(`  OK   ${nombre}`)
    return
  }
  fallos++
  console.log(`  FAIL ${nombre}${detalle ? ` -> ${detalle}` : ''}`)
}

function near(a, b, tol = 0.01) {
  return Math.abs(Number(a ?? 0) - Number(b ?? 0)) <= tol
}

function libroConFilas(nombreHoja, encabezados, filas) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet(nombreHoja)
  ws.addRow(encabezados)
  for (const fila of filas) ws.addRow(fila)
  return wb
}

// ---------------------------------------------------------------------------
// 1. Detalle por concepto: mes, catálogo, duplicados y desconocidos
// ---------------------------------------------------------------------------
function casoDetallePorConcepto() {
  console.log('\n1) Detalle por concepto (largo)')
  const wb = libroConFilas(
    'Carga agrupada',
    ['NRO. DOC. IDENTIDAD', 'Apellido Paterno', 'Nombres', 'PERIODO', 'CONCEPTO', 'IMPORTE', 'MEDIDA'],
    [
      ['12345678', 'Perez', 'Ana', '2025-01', 'SUELDO', 1000, 'DEVENGADO'],
      ['12345678', 'Perez', 'Ana', '2025-01', 'ASIG. FAM.', 100, 'DEVENGADO'],
      ['12345678', 'Perez', 'Ana', '2025-01', 'SUELDO', 1000, 'DEVENGADO'],
      ['12345678', 'Perez', 'Ana', '2025-01', 'AFP', 50, 'DEVENGADO'],
      ['12345678', 'Perez', 'Ana', '2025-02', 'BONIFICACION', 200, 'DEVENGADO'],
      ['12345678', 'Perez', 'Ana', '2025-02', 'CONCEPTO NUEVO', 300, 'DEVENGADO'],
      ['87654321', 'Lopez', 'Luis', '2025-02', 'HORAS EXTRA', 150, 'DEVENGADO'],
    ]
  )

  const r = motor.extraerRemuneraciones(wb, [], { loteId: 'largo' })
  const ana = r.importados.find((w) => w.dni === '12345678')
  const luis = r.importados.find((w) => w.dni === '87654321')

  check('modo atomico', r.modo === 'atomico', r.modo)
  check('dos trabajadores', r.importados.length === 2, String(r.importados.length))
  check('total sin duplicar ni desconocidos', near(ana.remuneraciones.total, 1300), ana.remuneraciones.total)
  check('enero = sueldo + familiar', near(ana.remuneraciones.enero, 1100), ana.remuneraciones.enero)
  check('febrero = bonificacion', near(ana.remuneraciones.febrero, 200), ana.remuneraciones.febrero)
  check('duplicado detectado', ana.remuneracionResumen.duplicados === 1, String(ana.remuneracionResumen.duplicados))
  check('descuento excluido', ana.remuneracionResumen.conceptosExcluidos.includes('Descuentos y aportes'), JSON.stringify(ana.remuneracionResumen.conceptosExcluidos))
  check('concepto desconocido listado', ana.remuneracionResumen.conceptosDesconocidos.includes('CONCEPTO NUEVO'), JSON.stringify(ana.remuneracionResumen.conceptosDesconocidos))
  check('sin conflictos', ana.remuneracionResumen.conflictos.length === 0, JSON.stringify(ana.remuneracionResumen.conflictos))
  check('segundo trabajador', near(luis.remuneraciones.total, 150), luis.remuneraciones.total)
  check('dias no derivados de remuneracion', r.importados.every((w) => w.diasTrabajados === undefined))
  check('lote registrado', ana.loteRemuneracion === 'largo', ana.loteRemuneracion)
  check('provenance con origen', ana.remuneracionResumen.provenance.every((p) => p.hoja && p.fila && p.rutaEncabezado), 'provenance incompleta')
}

// ---------------------------------------------------------------------------
// 2. Detalle + total declarado en la misma tabla
// ---------------------------------------------------------------------------
function casoDetalleConTotalDeclarado() {
  console.log('\n2) Detalle con total declarado (control de doble conteo)')
  const wb = libroConFilas(
    'Planilla',
    ['DNI', 'Nombres', 'PERIODO', 'CONCEPTO', 'IMPORTE', 'TOTAL DECLARADO'],
    [
      ['11111111', 'Ana Perez', '2025-01', 'SUELDO', 1000, 2500],
      ['11111111', 'Ana Perez', '2025-02', 'SUELDO', 1000, ''],
      ['11111111', 'Ana Perez', '2025-03', 'BONO', 500, ''],
    ]
  )

  const r = motor.extraerRemuneraciones(wb, [], { loteId: 'mixto' })
  const w = r.importados[0]
  const resumen = w.remuneracionResumen

  check('total = suma del detalle', near(w.remuneraciones.total, 2500), w.remuneraciones.total)
  check('total declarado no se suma dos veces', near(w.remuneraciones.total, resumen.totalDesdeHechos), JSON.stringify([w.remuneraciones.total, resumen.totalDesdeHechos]))
  check('base mixta', resumen.base === 'mixto', resumen.base)
  check('total declarado registrado', near(resumen.totalDeclarado, 2500), resumen.totalDeclarado)
  check('diferencia en cero', near(resumen.diferencia, 0), resumen.diferencia)
  check('sin conflictos', resumen.conflictos.length === 0, JSON.stringify(resumen.conflictos))
  check('meses disponibles', resumen.mensualDisponible === true)
  check('total con provenance', resumen.provenance.some((p) => p.origen === 'total'), 'sin provenance de total')
}

// ---------------------------------------------------------------------------
// 3. Total declarado que no cuadra con el detalle
// ---------------------------------------------------------------------------
function casoTotalDiscrepante() {
  console.log('\n3) Total declarado discrepante')
  const wb = libroConFilas(
    'Planilla',
    ['DNI', 'Nombres', 'PERIODO', 'CONCEPTO', 'IMPORTE', 'TOTAL DECLARADO'],
    [['22222222', 'Luis Lopez', '2025-01', 'SUELDO', 1000, 9999]],
  )

  const r = motor.extraerRemuneraciones(wb, [], { loteId: 'discrepante' })
  const resumen = r.importados[0].remuneracionResumen

  check('gana el detalle atomico', near(r.importados[0].remuneraciones.total, 1000), r.importados[0].remuneraciones.total)
  check('conflicto reportado', resumen.conflictos.length === 1, JSON.stringify(resumen.conflictos))
}

// ---------------------------------------------------------------------------
// 4. Catálogo configurable
// ---------------------------------------------------------------------------
function casoCatalogoConfigurable() {
  console.log('\n4) Catálogo configurable')
  const wb = libroConFilas(
    'Carga',
    ['DNI', 'Nombres', 'PERIODO', 'CONCEPTO', 'IMPORTE'],
    [
      ['33333333', 'Ana Perez', '2025-01', 'PAGO BASE', 1000],
      ['33333333', 'Ana Perez', '2025-01', 'BONO EXTRA', 100],
    ]
  )

  const porDefecto = motor.extraerRemuneraciones(wb, [], { loteId: 'cat' })
  check('concepto fuera de catálogo se excluye por defecto', near(porDefecto.importados[0].remuneraciones.total, 100), porDefecto.importados[0].remuneraciones.total)
  check('listado de desconocidos', porDefecto.importados[0].remuneracionResumen.conceptosDesconocidos.includes('PAGO BASE'), JSON.stringify(porDefecto.importados[0].remuneracionResumen.conceptosDesconocidos))

  const catalogo = atomico.crearCatalogoConceptos({
    remuneracion_basica: { aliases: ['SUELDO', 'PAGO BASE'] },
  })
  const personalizado = motor.extraerRemuneraciones(wb, [], { loteId: 'cat', catalogo })
  check('alias configurado se suma', near(personalizado.importados[0].remuneraciones.total, 1100), personalizado.importados[0].remuneraciones.total)
  check('sin desconocidos con catálogo propio', personalizado.importados[0].remuneracionResumen.conceptosDesconocidos.length === 0, JSON.stringify(personalizado.importados[0].remuneracionResumen.conceptosDesconocidos))
}

// ---------------------------------------------------------------------------
// 5. Solo total anual
// ---------------------------------------------------------------------------
function casoSoloTotal() {
  console.log('\n5) Solo total anual')
  const wb = libroConFilas('Remuneracion', ['DNI', 'Apellido Paterno', 'Nombres', 'TOTAL REMUNERACION'], [
    ['44444444', 'Perez', 'Ana', 50000],
  ])

  const r = motor.extraerRemuneraciones(wb, [], { loteId: 'total' })
  const w = r.importados[0]

  check('no bloquea la importacion', r.importados.length === 1, String(r.importados.length))
  check('usa el total declarado', near(w.remuneraciones.total, 50000), w.remuneraciones.total)
  check('base declarada', w.remuneracionResumen.base === 'declarado', w.remuneracionResumen.base)
  check('meses no disponibles', w.mesesRemuneracionDisponibles.length === 0, JSON.stringify(w.mesesRemuneracionDisponibles))
  check('sin falso conflicto', w.remuneracionResumen.conflictos.length === 0, JSON.stringify(w.remuneracionResumen.conflictos))
  check('queda marcado como incompleto', w.remuneracionIncompleta === true)
}

// ---------------------------------------------------------------------------
// 6. Lotes: reemplazo y trazabilidad
// ---------------------------------------------------------------------------
function casoLotes() {
  console.log('\n6) Reimportación por lotes')
  const nuevoLibro = () => libroConFilas('Remuneracion', ['DNI', 'Nombres', 'TOTAL REMUNERACION'], [['55555555', 'Ana Perez', 50000]])

  const primero = motor.extraerRemuneraciones(nuevoLibro(), [], { loteId: 'marzo.xlsx' })
  check('primer lote', primero.importados[0].loteRemuneracion === 'marzo.xlsx', primero.importados[0].loteRemuneracion)

  const segundo = motor.extraerRemuneraciones(nuevoLibro(), primero.trabajadores, { loteId: 'abril.xlsx' })
  check('lote nuevo reemplaza', segundo.trabajadores[0].loteRemuneracion === 'abril.xlsx', segundo.trabajadores[0].loteRemuneracion)
  check('lote anterior conservado', segundo.trabajadores[0].remuneracionResumen.lotesAnteriores.includes('marzo.xlsx'), JSON.stringify(segundo.trabajadores[0].remuneracionResumen.lotesAnteriores))

  const repetido = motor.extraerRemuneraciones(nuevoLibro(), segundo.trabajadores, { loteId: 'abril.xlsx' })
  check('mismo lote se reemplaza sin duplicarse', repetido.trabajadores.length === 1, String(repetido.trabajadores.length))
  check('mismo lote no figura como anterior', !repetido.trabajadores[0].remuneracionResumen.lotesAnteriores.includes('abril.xlsx'), JSON.stringify(repetido.trabajadores[0].remuneracionResumen.lotesAnteriores))
}

// ---------------------------------------------------------------------------
// 7. Días independientes de remuneración
// ---------------------------------------------------------------------------
function casoDiasIndependientes() {
  console.log('\n7) Días independientes')
  const wb = libroConFilas('Remuneracion', ['DNI', 'Nombres', 'ENERO', 'FEBRERO', 'TOTAL'], [['66666666', 'Ana Perez', 1000, 2000, 3000]])
  const rem = motor.extraerRemuneraciones(wb, [], { loteId: 'rem' })
  check('remuneracion sin dias', rem.importados.every((w) => w.diasTrabajados === undefined))

  const totales = new Map(rem.trabajadores.map((w) => [w.dni, w.remuneraciones.total]))
  const conDias = motor.extraerDias(wb, rem.trabajadores)
  const alterados = conDias.trabajadores.filter((w) => totales.has(w.dni) && !near(w.remuneraciones.total, totales.get(w.dni)))
  check('importar dias no altera remuneracion', alterados.length === 0, `${alterados.length} altered`)
}

// ---------------------------------------------------------------------------
// 8. Hoja dispersa de un millón de filas no debe colgar el escaneo
// ---------------------------------------------------------------------------
function casoHojaDispersa() {
  console.log('\n8) Hoja dispersa (1,048,576 filas declaradas)')
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('HOJA GRANDE')
  ws.getCell('A1').value = 'DNI'
  ws.getCell('B1').value = 'Nombres'
  ws.getCell('C1').value = 'TOTAL REMUNERACION'
  ws.getCell('A2').value = '77777777'
  ws.getCell('B2').value = 'Ana Perez'
  ws.getCell('C2').value = 12000
  ws.getCell('A1048576').value = 'sombra'

  const inicio = Date.now()
  const r = atomico.extraerRemuneracionAtomico(wb)
  const ms = Date.now() - inicio

  check('escaneo termina en menos de 30s', ms < 30000, `${ms}ms`)
  check('lee el dato real y no la fila fantasma', r.trabajadores.length === 1 && near(r.trabajadores[0].remuneraciones.total, 12000), JSON.stringify(r.trabajadores.map((w) => [w.dni, w.remuneraciones.total])))
}

async function main() {
  casoDetallePorConcepto()
  casoDetalleConTotalDeclarado()
  casoTotalDiscrepante()
  casoCatalogoConfigurable()
  casoSoloTotal()
  casoLotes()
  casoDiasIndependientes()
  casoHojaDispersa()

  console.log(fallos === 0 ? '\nTODAS LAS VERIFICACIONES ATOMICAS OK' : `\n${fallos} verificaciones fallidas`)
  process.exit(fallos === 0 ? 0 : 1)
}

main().catch((e) => {
  console.error(e)
  process.exit(1)
})
