import ExcelJS from 'exceljs'

type MonthlyValues = {
  enero: number
  febrero: number
  marzo: number
  abril: number
  mayo: number
  junio: number
  julio: number
  agosto: number
  septiembre: number
  octubre: number
  noviembre: number
  diciembre: number
  total: number
}

type Incidencia = {
  codigo: string
  descripcion: string
  dias: number
  diasNeto?: number
  fechaInicio: Date | null
  fechaFin: Date | null
}

type PeriodoLaboral = {
  inicio: string
  fin: string | null
  meses: number
  estado: 'Trabajó' | 'No trabajó' | 'Reingresó' | 'Cesó'
  descripcion?: string
}

type Reingreso = {
  fechaReingreso: string
  diasFuera: number
  periodoAnterior: string
  periodoActual: string
}

type TrabajadorCompleto = {
  dni: string
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string
  nombreCompleto: string
  diasPosibles: number
  diasNoLaboradosReferencia: number
  diasEfectivos: number
  diasNoLaborados: number
  diasCalculados: number
  fechaInicio: string | null
  fechaCese: string | null
  esIngreso2025: boolean
  periodos: PeriodoLaboral[]
  reingresos: Reingreso[]
  tieneReingreso: boolean
  incidencias: Incidencia[]
}

const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'] as const

const MESES_EN_CERO: MonthlyValues = {
  enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0,
  julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 0
}

const normalizeDni = (value: ExcelJS.CellValue): string | null => {
  const dni = String(value ?? '').trim().replace(/\.0$/, '')
  return /^\d{7,8}$/.test(dni) ? dni.padStart(8, '0') : null
}

const toNumber = (value: ExcelJS.CellValue): number => {
  if (typeof value === 'number') return value
  if (typeof value === 'object' && value !== null && 'result' in value) {
    return toNumber(value.result as ExcelJS.CellValue)
  }
  const numero = Number(String(value ?? '').replace(',', '.'))
  return Number.isFinite(numero) ? numero : 0
}

const limpiarNombre = (valor: ExcelJS.CellValue): string => {
  if (!valor) return ''
  if (valor instanceof Date) return ''
  let texto = String(valor).trim()
  texto = texto.replace(/\s+(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+[A-Z][a-z]{2}\s+\d{1,2}\s+\d{4}.*$/gi, '')
  return texto.trim()
}

const formatDate = (date: Date): string => {
  const day = String(date.getUTCDate()).padStart(2, '0')
  const month = String(date.getUTCMonth() + 1).padStart(2, '0')
  const year = date.getUTCFullYear()
  return `${year}-${month}-${day}`
}

const parseExcelDate = (value: ExcelJS.CellValue): Date | null => {
  if (!value) return null
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value
  if (typeof value === 'number') {
    const excelEpoch = Date.UTC(1899, 11, 30)
    return new Date(excelEpoch + value * 86_400_000)
  }
  if (typeof value === 'string') {
    const date = new Date(value)
    return Number.isNaN(date.getTime()) ? null : date
  }
  if (typeof value === 'object' && value !== null) {
    if ('result' in value) return parseExcelDate(value.result as ExcelJS.CellValue)
    if ('text' in value && typeof value.text === 'string') return parseExcelDate(value.text)
  }
  return null
}

const diffDays = (start: Date, end: Date): number => {
  const diffMs = end.getTime() - start.getTime()
  return Math.round(diffMs / 86_400_000)
}

const calcularMeses = (inicio: Date, fin: Date): number => {
  const meses = (fin.getFullYear() - inicio.getFullYear()) * 12 + (fin.getMonth() - inicio.getMonth())
  return Math.max(1, meses + 1)
}

function extraerDiasEfectivos(hoja: ExcelJS.Worksheet): Map<string, { total: number; posibles: number; noLaboradosRef: number; nombre: string; fechaInicio: Date | null; fechaCese: Date | null }> {
  const resultado = new Map<string, { total: number; posibles: number; noLaboradosRef: number; nombre: string; fechaInicio: Date | null; fechaCese: Date | null }>()

  hoja.eachRow((row, rowNumber) => {
    if (rowNumber < 8) return

    const dni = normalizeDni(row.getCell(2).value)
    if (!dni) return

    const nombreCompleto = limpiarNombre(row.getCell(3).value)
    // Columna L = días posibles, columna M = deducciones, columna O = días efectivos (neto)
    const posibles = toNumber(row.getCell(12).value) || toNumber(row.getCell(7).value)
    const noLaboradosRef = toNumber(row.getCell(13).value)
    const total = toNumber(row.getCell(15).value) || toNumber(row.getCell(7).value)
    const fechaInicio = parseExcelDate(row.getCell(4).value)
    const fechaCese = parseExcelDate(row.getCell(6).value)

    resultado.set(dni, { total, posibles, noLaboradosRef, nombre: nombreCompleto, fechaInicio, fechaCese })
  })

  return resultado
}

function extraerIncidencias(hoja: ExcelJS.Worksheet): Map<string, { incidencias: Incidencia[]; info: { apellidoPaterno: string; apellidoMaterno: string; nombres: string } }> {
  const resultado = new Map<string, { incidencias: Incidencia[]; info: { apellidoPaterno: string; apellidoMaterno: string; nombres: string } }>()

  hoja.eachRow((row, rowNumber) => {
    if (rowNumber < 6) return

    const dni = normalizeDni(row.getCell(3).value)
    if (!dni) return

    const apellidoPaterno = limpiarNombre(row.getCell(4).value)
    const apellidoMaterno = limpiarNombre(row.getCell(5).value)
    const nombres = limpiarNombre(row.getCell(6).value)

    const descripcion = String(row.getCell(8).value ?? '').trim()
    const cantidadDias = toNumber(row.getCell(11).value)

    if (descripcion === '[object Object]' || descripcion === 'NaN' || (!descripcion && cantidadDias === 0)) {
      return
    }

    const existente = resultado.get(dni) ?? {
      incidencias: [],
      info: { apellidoPaterno: '', apellidoMaterno: '', nombres: '' }
    }

    existente.incidencias.push({
      codigo: String(row.getCell(7).value ?? '').trim(),
      descripcion,
      dias: cantidadDias,
      diasNeto: toNumber(row.getCell(14).value) || undefined,
      fechaInicio: parseExcelDate(row.getCell(15).value),
      fechaFin: parseExcelDate(row.getCell(16).value),
    })

    existente.info = {
      apellidoPaterno: apellidoPaterno || existente.info.apellidoPaterno,
      apellidoMaterno: apellidoMaterno || existente.info.apellidoMaterno,
      nombres: nombres || existente.info.nombres,
    }

    resultado.set(dni, existente)
  })

  return resultado
}

function calcularPeriodosLaborales(
  diasEfectivosTotal: number,
  incidencias: Incidencia[],
  fechaInicio?: Date | null,
  fechaCese?: Date | null
): { periodos: PeriodoLaboral[]; reingresos: Reingreso[]; esIngreso2025: boolean } {
  const periodos: PeriodoLaboral[] = []
  const inicio2025 = new Date('2025-01-01')
  const fin2025 = new Date('2025-12-31')

  // Determinar si el trabajador ingresó durante 2025
  const esIngreso2025 = fechaInicio ? fechaInicio.getFullYear() === 2025 : false

  // Si el trabajador empezó a mitad de año (ingreso 2025), el período
  // comienza desde su fecha de ingreso, no desde enero.
  const inicioEfectivo = esIngreso2025 && fechaInicio && fechaInicio > inicio2025
    ? fechaInicio
    : inicio2025

  if (incidencias.length === 0) {
    // Sin incidencias: el trabajador trabajó desde su inicio hasta fin de año (o cese)
    const finPeriodo = fechaCese && fechaCese < fin2025 ? fechaCese : fin2025
    const estado = esIngreso2025
      ? 'Trabajó'
      : (fechaCese ? 'Cesó' : 'Trabajó')

    periodos.push({
      inicio: formatDate(inicioEfectivo),
      fin: formatDate(finPeriodo),
      meses: calcularMeses(inicioEfectivo, finPeriodo),
      estado,
    })
    return { periodos, reingresos: [], esIngreso2025 }
  }

  const fechasOrdenadas = incidencias
    .filter(i => i.fechaInicio && i.fechaFin)
    .sort((a, b) => (a.fechaInicio?.getTime() ?? 0) - (b.fechaInicio?.getTime() ?? 0))

  let fechaActual = inicioEfectivo

  for (const incidencia of fechasOrdenadas) {
    if (!incidencia.fechaInicio || !incidencia.fechaFin) continue

    if (incidencia.fechaInicio > fechaActual) {
      periodos.push({
        inicio: formatDate(fechaActual),
        fin: formatDate(new Date(incidencia.fechaInicio.getTime() - 86_400_000)),
        meses: calcularMeses(fechaActual, new Date(incidencia.fechaInicio.getTime() - 86_400_000)),
        estado: 'Trabajó',
      })
    }

    periodos.push({
      inicio: formatDate(incidencia.fechaInicio),
      fin: formatDate(incidencia.fechaFin),
      meses: calcularMeses(incidencia.fechaInicio, incidencia.fechaFin),
      estado: 'No trabajó',
      descripcion: incidencia.descripcion,
    })

    fechaActual = new Date(incidencia.fechaFin.getTime() + 86_400_000)
  }

  if (fechaActual <= fin2025) {
    periodos.push({
      inicio: formatDate(fechaActual),
      fin: '2025-12-31',
      meses: calcularMeses(fechaActual, fin2025),
      estado: 'Trabajó',
    })
  }

  const reingresos: Reingreso[] = []
  let prevTrabajo = false
  let prevFin: Date | null = null

  for (const periodo of periodos) {
    if (periodo.estado === 'Trabajó') {
      if (prevTrabajo && prevFin) {
        const inicioActual = new Date(periodo.inicio)
        const diasFuera = diffDays(prevFin, inicioActual)
        // Solo considerar reingreso si el gap es significativo (más de 15 días),
        // de lo contrario es una licencia corta normal, no un reingreso.
        if (inicioActual > prevFin && diasFuera > 15) {
          reingresos.push({
            fechaReingreso: periodo.inicio,
            diasFuera,
            periodoAnterior: formatDate(prevFin),
            periodoActual: periodo.inicio,
          })
          periodo.estado = 'Reingresó'
        }
      }
      prevTrabajo = true
      prevFin = new Date(periodo.fin ?? '2025-12-31')
    } else {
      prevTrabajo = false
      if (periodo.fin) {
        prevFin = new Date(periodo.fin)
      }
    }
  }

  return { periodos, reingresos, esIngreso2025 }
}

async function main() {
  const archivo = process.argv[2] || 'PLANILLA UTILIDADES 2025 - CITIKOLD.xlsx'

  console.log('========================================')
  console.log('EXTRACCIÓN DE DÍAS - FORMATO CITIKOLD')
  console.log('========================================')
  console.log(`Archivo: ${archivo}`)

  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(archivo)

  console.log('\nHojas encontradas:')
  workbook.worksheets.forEach(sheet => console.log(`  - ${sheet.name}`))

  const hojaEfectivos = workbook.getWorksheet('DÍAS EFECTIVOS LABORADOS 2025')
  const hojaIncidencias = workbook.getWorksheet('DIAS NO LABORADOS')

  if (!hojaEfectivos) {
    console.error('\n❌ No se encontró la hoja "DÍAS EFECTIVOS LABORADOS 2025"')
    return
  }

  if (!hojaIncidencias) {
    console.error('\n❌ No se encontró la hoja "DIAS NO LABORADOS"')
    return
  }

  console.log('\n✅ Hojas requeridas encontradas')

  console.log('\n--- Extrayendo días efectivos ---')
  const diasEfectivos = extraerDiasEfectivos(hojaEfectivos)
  console.log(`Trabajadores encontrados: ${diasEfectivos.size}`)

  console.log('\n--- Extrayendo incidencias ---')
  const incidenciasData = extraerIncidencias(hojaIncidencias)
  console.log(`Trabajadores con incidencias: ${incidenciasData.size}`)

  const trabajadores: TrabajadorCompleto[] = []

  for (const [dni, diasData] of diasEfectivos) {
    const incidenciaData = incidenciasData.get(dni)
    const incidencias = incidenciaData?.incidencias ?? []

    // Días no laborados: prioridad a la referencia del cuadro (columna M),
    // que suma los "DIAS EFECTIVOS TOTAL" (columna N) de cada incidencia.
    const diasNoLaborados =
      diasData.noLaboradosRef > 0
        ? diasData.noLaboradosRef
        : incidencias.reduce((sum, inc) => sum + (inc.diasNeto ?? 0), 0)
    // Días efectivos calculados = días posibles − días no laborados (un solo descuento).
    // Debe coincidir con la columna O (días efectivos del Excel).
    const diasCalculados = Math.max(0, diasData.posibles - diasNoLaborados)

    const { periodos, reingresos, esIngreso2025 } = calcularPeriodosLaborales(
      diasData.total,
      incidencias,
      diasData.fechaInicio,
      diasData.fechaCese
    )

    const info = incidenciaData?.info
    const partes = diasData.nombre.split(/\s+/).filter(Boolean)

    trabajadores.push({
      dni,
      apellidoPaterno: info?.apellidoPaterno || partes[0] || '',
      apellidoMaterno: info?.apellidoMaterno || partes[1] || '',
      nombres: info?.nombres || partes.slice(2).join(' ') || '',
      nombreCompleto: diasData.nombre,
      diasPosibles: diasData.posibles,
      diasNoLaboradosReferencia: diasData.noLaboradosRef,
      diasEfectivos: diasData.total,
      diasNoLaborados,
      diasCalculados,
      fechaInicio: diasData.fechaInicio ? formatDate(diasData.fechaInicio) : null,
      fechaCese: diasData.fechaCese && !Number.isNaN(diasData.fechaCese.getTime()) ? formatDate(diasData.fechaCese) : null,
      esIngreso2025,
      periodos,
      reingresos,
      tieneReingreso: reingresos.length > 0,
      incidencias,
    })

    incidenciasData.delete(dni)
  }

  for (const [dni, incidenciaData] of incidenciasData) {
    const diasNoLaborados = incidenciaData.incidencias.reduce((sum, inc) => sum + (inc.diasNeto ?? inc.dias), 0)

    trabajadores.push({
      dni,
      apellidoPaterno: incidenciaData.info.apellidoPaterno,
      apellidoMaterno: incidenciaData.info.apellidoMaterno,
      nombres: incidenciaData.info.nombres,
      nombreCompleto: `${incidenciaData.info.apellidoPaterno} ${incidenciaData.info.apellidoMaterno} ${incidenciaData.info.nombres}`.trim(),
      diasPosibles: 0,
      diasNoLaboradosReferencia: diasNoLaborados,
      diasEfectivos: 0,
      diasNoLaborados,
      diasCalculados: 0,
      fechaInicio: null,
      fechaCese: null,
      esIngreso2025: false,
      periodos: [],
      reingresos: [],
      tieneReingreso: false,
      incidencias: incidenciaData.incidencias,
    })
  }

  trabajadores.sort((a, b) => a.apellidoPaterno.localeCompare(b.apellidoPaterno))

  console.log('\n========================================')
  console.log('RESUMEN DE EXTRACCIÓN')
  console.log('========================================')
  console.log(`Total trabajadores: ${trabajadores.length}`)
  console.log(`Con días efectivos: ${trabajadores.filter(t => t.diasEfectivos > 0).length}`)
  console.log(`Con incidencias: ${trabajadores.filter(t => t.incidencias.length > 0).length}`)
  console.log(`Con reingresos: ${trabajadores.filter(t => t.tieneReingreso).length}`)

  console.log('\n========================================')
  console.log('DETALLE POR TRABAJADOR')
  console.log('========================================')

  for (const t of trabajadores) {
    console.log(`\n--- ${t.nombreCompleto} (DNI: ${t.dni}) ---`)
    if (t.fechaInicio) console.log(`  Fecha de inicio: ${t.fechaInicio} ${t.esIngreso2025 ? '(INGRESO 2025)' : ''}`)
    if (t.fechaCese) console.log(`  Fecha de cese: ${t.fechaCese}`)
    console.log(`  Días posibles: ${t.diasPosibles}`)
    console.log(`  Días no laborados: ${t.diasNoLaborados} (referencia del cuadro: ${t.diasNoLaboradosReferencia})`)
    console.log(`  Días efectivos (hoja): ${t.diasEfectivos}`)
    console.log(`  Días calculados (posibles − no laborados): ${t.diasCalculados}`)

    if (t.incidencias.length > 0) {
      console.log(`  Incidencias (${t.incidencias.length}):`)
      for (const inc of t.incidencias) {
        console.log(`    - ${inc.descripcion}: ${inc.dias} días`)
        if (inc.fechaInicio && inc.fechaFin) {
          console.log(`      Del ${formatDate(inc.fechaInicio)} al ${formatDate(inc.fechaFin)}`)
        }
      }
    }

    if (t.periodos.length > 0) {
      console.log(`  Períodos laborales (${t.periodos.length}):`)
      for (const p of t.periodos) {
        console.log(`    - ${p.estado}: ${p.inicio} al ${p.fin ?? 'presente'} (${p.meses} meses)`)
        if (p.descripcion) {
          console.log(`      ${p.descripcion}`)
        }
      }
    }

    if (t.reingresos.length > 0) {
      console.log(`  REINGRESOS (${t.reingresos.length}):`)
      for (const r of t.reingresos) {
        console.log(`    - Reingresó el ${r.fechaReingreso}`)
        console.log(`      estuvo ${r.diasFuera} días fuera`)
      }
    }
  }

  const output = {
    fechaExtraccion: new Date().toISOString(),
    archivo,
    resumen: {
      totalTrabajadores: trabajadores.length,
      conDiasEfectivos: trabajadores.filter(t => t.diasEfectivos > 0).length,
      conIncidencias: trabajadores.filter(t => t.incidencias.length > 0).length,
      conReingresos: trabajadores.filter(t => t.tieneReingreso).length,
      ingresos2025: trabajadores.filter(t => t.esIngreso2025).length,
      totalDiasEfectivos: trabajadores.reduce((sum, t) => sum + t.diasEfectivos, 0),
      totalDiasNoLaborados: trabajadores.reduce((sum, t) => sum + t.diasNoLaborados, 0),
    },
    trabajadores,
  }

  const fs = await import('fs')
  const outputPath = 'resultado-extraccion-citikold.json'
  fs.writeFileSync(outputPath, JSON.stringify(output, null, 2))
  console.log(`\n✅ Resultado guardado en: ${outputPath}`)
}

main().catch(console.error)