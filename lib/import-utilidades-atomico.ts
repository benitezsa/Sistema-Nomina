import * as ExcelJS from 'exceljs'

export type MonthlyValues = {
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

const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'] as const
const MESES_EN_CERO: MonthlyValues = {
  enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0,
  julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0,
  total: 0,
}

export type MedidaRemuneracion = 'bruto' | 'devengado' | 'pagado' | 'neto'
export type RowKind = 'detalle' | 'total' | 'subtotal' | 'footer'
export type ConceptFamily = 'remunerativo' | 'no_remunerativo' | 'agregado' | 'desconocido'
export type EstadoRegistro = 'considerado' | 'excluido' | 'duplicado' | 'desconocido'

export type ConceptClassification = { key: string; label: string; family: ConceptFamily }

export type DefinicionConcepto = {
  key: string
  label: string
  family: Exclude<ConceptFamily, 'desconocido'>
  aliases: string[]
}

export type CatalogoConceptos = Record<string, DefinicionConcepto>

export type PeriodoAtomico = {
  key: string
  label: string
  mes?: string
  anio?: number
}

export type Provenance = {
  loteId: string
  hoja: string
  fila: number
  columna?: number
  rutaEncabezado: string[]
  valorOriginal: string
  periodo?: string
  origen?: string
  concepto?: string
}

export type RegistroRemuneracion = {
  id: string
  loteId: string
  dni: string
  periodo: PeriodoAtomico | null
  conceptKey: string
  conceptLabel: string
  medida: MedidaRemuneracion | 'total'
  monto: number
  moneda: 'PEN'
  rowKind: RowKind
  estado: EstadoRegistro
  origen: 'concepto' | 'mensual' | 'total'
  provenance: Provenance
}

export type ResumenRemuneracion = {
  loteId: string
  lotesAnteriores: string[]
  base: 'detalle' | 'declarado' | 'mixto'
  totalDesdeHechos: number
  totalDeclarado?: number
  diferencia?: number
  mensualDisponible: boolean
  considerados: number
  excluidos: number
  duplicados: number
  conflictos: string[]
  conceptosDesconocidos: string[]
  conceptosExcluidos: string[]
  conceptos: Record<string, number>
  provenance: Provenance[]
}

export type ProyeccionRemuneracion = {
  dni: string
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string
  fechaInicio: string
  fechaCese: string
  remuneraciones: MonthlyValues
  remuneracionIncompleta: boolean
  mesesRemuneracionDisponibles: string[]
  mesesRemuneraciones: string[]
  mesesPresentes: string[]
  remuneracionResumen: ResumenRemuneracion
}

export type OpcionesRemuneracionAtomica = {
  loteId?: string
  lotesAnteriores?: string[]
  catalogo?: CatalogoConceptos
}

export type ResultadoRemuneracionAtomica = {
  disponible: boolean
  modo: 'detalle' | 'mensual' | 'declarado' | 'indisponible'
  trabajadores: ProyeccionRemuneracion[]
  registros: RegistroRemuneracion[]
}

const DEFINICIONES_POR_DEFECTO: DefinicionConcepto[] = [
  {
    key: 'remuneracion_basica',
    label: 'Remuneración básica',
    family: 'remunerativo',
    aliases: ['SUELDO', 'SUELDO BASICO', 'REMUNERACION BASICA', 'REM. BASICA', 'REM BASICA', 'REMUN. BASICA', 'REMUN BASICA', 'S. BASICO', 'SALARIO BASE', 'REMUNERACION'],
  },
  {
    key: 'asignacion_familiar',
    label: 'Asignación familiar',
    family: 'remunerativo',
    aliases: ['ASIG FAM', 'ASIG. FAM', 'ASIGNACION FAMILIAR', 'ASIGNACION FAMILIAR', 'ALIMENTACION FAMILIAR'],
  },
  {
    key: 'bonificacion',
    label: 'Bonificación',
    family: 'remunerativo',
    aliases: ['BONIFICACION', 'BONO', 'BONIFICACION EXTRA', 'BONO EXTRA'],
  },
  {
    key: 'gratificacion',
    label: 'Gratificación',
    family: 'remunerativo',
    aliases: ['GRATIFICACION', 'AGUINALDO'],
  },
  {
    key: 'horas_extra',
    label: 'Horas extra',
    family: 'remunerativo',
    aliases: ['HORAS EXTRA', 'HORAS EXTRAS', 'HORA EXTRA', 'HORAS EXTRAS 100%'],
  },
  {
    key: 'movilidad',
    label: 'Movilidad',
    family: 'remunerativo',
    aliases: ['MOVILIDAD', 'MOV SUPEDITADA', 'MOV. SUPEDITADA', 'MOVILIDAD LIBRE DISPOSICION'],
  },
  {
    key: 'refrigerio',
    label: 'Refrigerio',
    family: 'remunerativo',
    aliases: ['REFRIGERIO', 'ALIMENTACION', 'REFRIGERIO SUPEDITADO', 'REFRIGERIO FIJO'],
  },
  {
    key: 'utilidad_convencional',
    label: 'Utilidad convencional',
    family: 'remunerativo',
    aliases: ['UTILIDAD CONVENCIONAL', 'UTILIDADES'],
  },
  {
    key: 'asignacion_festividad',
    label: 'Asignación por festividad',
    family: 'remunerativo',
    aliases: ['ASIGNACION POR FESTIVIDAD', 'BONO ANUAL', 'BONO FESTIVIDAD'],
  },
  {
    key: 'remuneracion_mensual',
    label: 'Remuneración mensual',
    family: 'remunerativo',
    aliases: ['REMUNERACION MENSUAL', 'MENSUAL'],
  },
  {
    key: 'descuentos',
    label: 'Descuentos y aportes',
    family: 'no_remunerativo',
    aliases: ['AFP', 'ONP', 'ESSALUD', '5TA CATEGORIA', 'IMPUESTO', 'RETENCION', 'PRESTAMO', 'ADELANTO', 'DESCUENTO', 'CTS', 'NETO A PAGAR', 'TOTAL DESCUENTOS', 'TOTAL APORTES'],
  },
]

export function crearCatalogoConceptos(overrides: Record<string, Partial<DefinicionConcepto>> = {}): CatalogoConceptos {
  const catalogo: CatalogoConceptos = {}
  for (const definicion of DEFINICIONES_POR_DEFECTO) {
    catalogo[definicion.key] = { ...definicion, aliases: [...definicion.aliases] }
  }

  for (const [key, patch] of Object.entries(overrides)) {
    const actual = catalogo[key]
    catalogo[key] = {
      key,
      label: patch.label ?? actual?.label ?? key,
      family: patch.family ?? actual?.family ?? 'no_remunerativo',
      aliases: patch.aliases ?? actual?.aliases ?? [patch.label ?? key],
    }
  }

  return catalogo
}

export const CATALOGO_CONCEPTOS_POR_DEFECTO = crearCatalogoConceptos()

function normalizeText(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''

  if (typeof value === 'object') {
    if ('text' in value) return String(value.text ?? '')
    if ('result' in value) return String(value.result ?? '')
  }

  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase()
}

function compactText(value: string): string {
  return normalizeText(value).replace(/[^A-Z0-9]/g, '')
}

function normalizeDni(value: ExcelJS.CellValue): string | null {
  const dni = String(value ?? '').trim().replace(/\.0$/, '')
  if (/^\d{9}$/.test(dni)) return dni
  return /^\d{7,8}$/.test(dni) ? dni.padStart(8, '0') : null
}

function toNumberNullable(value: ExcelJS.CellValue): number | null {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'object') {
    if ('result' in value) return toNumberNullable(value.result as ExcelJS.CellValue)
    if ('text' in value) return toNumberNullable(value.text)
    return null
  }

  const texto = String(value).trim()
  if (!texto) return null
  const numero = Number(texto.replace(/[^\d,.-]/g, '').replace(',', '.'))
  return Number.isFinite(numero) ? numero : null
}

const MESES_POR_TEXTO: Record<string, string> = {
  ENERO: 'enero',
  ENE: 'enero',
  FEBRERO: 'febrero',
  FEB: 'febrero',
  MARZO: 'marzo',
  MAR: 'marzo',
  ABRIL: 'abril',
  ABR: 'abril',
  MAYO: 'mayo',
  MAY: 'mayo',
  JUNIO: 'junio',
  JUN: 'junio',
  JULIO: 'julio',
  JUL: 'julio',
  AGOSTO: 'agosto',
  AGO: 'agosto',
  SETIEMBRE: 'septiembre',
  SEPTIEMBRE: 'septiembre',
  SEP: 'septiembre',
  SET: 'septiembre',
  OCTUBRE: 'octubre',
  OCT: 'octubre',
  NOVIEMBRE: 'noviembre',
  NOV: 'noviembre',
  DICIEMBRE: 'diciembre',
  DIC: 'diciembre',
}

function excelDateFromSerial(serial: number): Date | null {
  if (!Number.isFinite(serial) || serial < 20000 || serial > 80000) return null
  const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000)
  return Number.isNaN(date.getTime()) ? null : date
}

function periodFromValue(value: ExcelJS.CellValue): PeriodoAtomico | null {
  if (value === null || value === undefined || value === '') return null

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null
    const mes = MONTHS[value.getUTCMonth()]
    return { key: `${value.getUTCFullYear()}-${mes}`, label: `${value.getUTCFullYear()}-${mes}`, mes, anio: value.getUTCFullYear() }
  }

  if (typeof value === 'number') {
    const date = excelDateFromSerial(value)
    return date ? periodFromValue(date) : null
  }

  if (typeof value === 'object') {
    if ('result' in value) return periodFromValue(value.result as ExcelJS.CellValue)
    if ('text' in value) return periodFromValue(value.text)
    return null
  }

  const texto = String(value).trim()
  if (!texto) return null
  const isoMes = /^(\d{4})[-/](\d{1,2})/.exec(texto)
  if (isoMes) {
    const anio = Number(isoMes[1])
    const indice = Number(isoMes[2]) - 1
    if (indice >= 0 && indice < MONTHS.length) {
      const mes = MONTHS[indice]
      return { key: `${anio}-${mes}`, label: `${anio}-${mes}`, mes, anio }
    }
  }

  const mesAnio = /^(\d{1,2})[-/](\d{4})/.exec(texto)
  if (mesAnio) {
    const indice = Number(mesAnio[1]) - 1
    const anio = Number(mesAnio[2])
    if (indice >= 0 && indice < MONTHS.length) {
      const mes = MONTHS[indice]
      return { key: `${anio}-${mes}`, label: `${anio}-${mes}`, mes, anio }
    }
  }

  const normalizado = normalizeText(texto)
  const mesTexto = normalizado.split(/\s+/).map((parte) => MESES_POR_TEXTO[parte]).find(Boolean)
  if (mesTexto) {
    const anio = /(?:^|\D)(20\d{2})(?:\D|$)/.exec(normalizado)?.[1]
    return {
      key: anio ? `${anio}-${mesTexto}` : `mes:${mesTexto}`,
      label: anio ? `${anio}-${mesTexto}` : mesTexto.toUpperCase(),
      mes: mesTexto,
      ...(anio ? { anio: Number(anio) } : {}),
    }
  }

  return null
}

function periodFromHeader(path: string[]): PeriodoAtomico | null {
  for (const segment of [...path].reverse()) {
    const period = periodFromValue(segment)
    if (period) return period
  }
  return null
}

function isPeriodHeader(path: string[]): boolean {
  const text = normalizeText(path.join(' '))
  return /(^|\s)(MES|PERIODO|PLANILLA|CICLO)(\s|$)/.test(text) && !/FECHA (INGRESO|CESE|NACIMIENTO)/.test(text)
}

function isNameHeader(path: string[], names: string[]): boolean {
  return path.some((segment) => names.includes(normalizeText(segment)))
}

function isExplicitConceptHeader(path: string[]): boolean {
  return path.some((segment) => /(^|\s)(CONCEPTO|RUBRO|ITEM|DESCRIPCION|DETALLE)(\s|$)/.test(normalizeText(segment)))
}

function isAmountHeader(path: string[], classification: { family: ConceptFamily }): boolean {
  const text = normalizeText(path.join(' '))
  if (classification.family === 'agregado') {
    if (/REMUN|INGRESO|SUELDO|DEVENGAD|PAGAD|BRUTO|NETO/.test(text)) return true
    return /DECLARADO|IMPORTE|MONTO|VALOR/.test(text)
  }
  if (classification.family !== 'desconocido') return true
  return /MONTO|IMPORTE|VALOR|SUELDO|REMUNER/.test(text) && !/FECHA|DOCUMENTO|CODIGO|CONCEPTO|DESCRIPCION|NRO|NUMERO DE CUENTA|INTERBANCARIO|CUENTA/.test(text)
}

function isAggregateHeader(path: string[]): boolean {
  return /TOTAL|SUBTOTAL|RESUMEN/.test(normalizeText(path.join(' ')))
}

function measureFromText(value: ExcelJS.CellValue): MedidaRemuneracion | 'total' | null {
  const text = normalizeText(value)
  if (!text) return null
  if (/NETO/.test(text)) return 'neto'
  if (/DEVENGAD/.test(text)) return 'devengado'
  if (/PAGAD/.test(text)) return 'pagado'
  if (/BRUTO/.test(text)) return 'bruto'
  if (/TOTAL|SUBTOTAL/.test(text)) return 'total'
  return null
}

function currencyIsPen(value: string): boolean {
  return !/DOLAR|USD|EURO|EUR/.test(normalizeText(value))
}

function classificationFor(label: string, catalogo: CatalogoConceptos): { key: string; label: string; family: ConceptFamily } {
  const normalized = normalizeText(label)
  const compact = compactText(normalized)
  if (!compact) return { key: 'desconocido:', label: normalized, family: 'desconocido' }
  if (/TOTAL|SUBTOTAL|RESUMEN/.test(normalized)) return { key: 'agregado', label: normalized, family: 'agregado' }

  for (const definicion of Object.values(catalogo)) {
    if (definicion.aliases.some((alias) => compactText(alias) === compact)) {
      return { key: definicion.key, label: definicion.label, family: definicion.family }
    }
  }

  return { key: `desconocido:${compact}`, label: normalized, family: 'desconocido' }
}

function classificationFromPath(path: string[], catalogo: CatalogoConceptos): { key: string; label: string; family: ConceptFamily } {
  for (const segment of [...path].reverse()) {
    const classification = classificationFor(segment, catalogo)
    if (classification.family !== 'desconocido' && classification.family !== 'agregado') return classification
  }
  return classificationFor(path[path.length - 1] ?? '', catalogo)
}

function documentScore(path: string[]): number {
  const compact = compactText(path.join(' '))
  if (/(^|\s)(TIPO|TYPE)(\s|$)/.test(normalizeText(path.join(' ')))) return 0
  if (compact === 'DNI') return 100
  if (['NRODOCIDE', 'NRODOCIDENTIDAD', 'NUMDOCIDE', 'NUMERODEDOCUMENTO', 'DOCUMENTODEIDENTIDAD', 'NUMERODEIDENTIDAD', 'NDOCUMENTO', 'NRODOCUMENTO'].includes(compact)) return 90
  if (compact === 'NUMERO' || compact === 'NRO') return 30
  const text = normalizeText(path.join(' '))
  const tieneDocumento = /DOCUMENTO|DOCUMENT|DOC/.test(text)
  const tieneNumero = /(^|\s)(NRO|NUMERO|NUM|N)(\s|$)/.test(text) || compact.startsWith('NRO') || compact.startsWith('NUM')
  return tieneDocumento && (tieneNumero || /IDENT/.test(text)) ? 80 : 0
}

type CandidateTable = {
  ws: ExcelJS.Worksheet
  headerRows: number[]
  dataRow: number
  dniCol: number
  paths: Map<number, string[]>
}

type ColumnInfo = {
  col: number
  path: string[]
  label: string
  identity: boolean
  nameApellido: boolean
  nameMaterno: boolean
  nameNombres: boolean
  nameCompleto: boolean
  period: boolean
  explicitConcept: boolean
  knownConcept: boolean
  aggregate: boolean
  amount: boolean
  month: PeriodoAtomico | null
  measure: MedidaRemuneracion | 'total' | null
}

function nameKind(path: string[]): 'apellido' | 'materno' | 'nombres' | 'completo' | null {
  const text = normalizeText(path.join(' '))
  if (/APELLIDO.*PATERNO|APE.*PATERNO/.test(text)) return 'apellido'
  if (/APELLIDO.*MATERNO|APE.*MATERNO/.test(text)) return 'materno'
  if (/APELLIDOS.*NOMBRES|NOMBRES.*APELLIDOS/.test(text)) return 'completo'
  if (/(^|\s)NOMBRES?(\s|$)/.test(text)) return 'nombres'
  return null
}

function dateColumn(path: string[], kind: 'ingreso' | 'cese'): boolean {
  const text = normalizeText(path.join(' '))
  if (kind === 'ingreso') return /FECHA.*(INGRESO|INICIO)|FECHA INGRESO|FECHA DE INICIO/.test(text)
  return /FECHA.*(CESE|FIN)|FECHA CESE|FECHA DE FIN/.test(text)
}

function dateFromValue(value: ExcelJS.CellValue): Date | null {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value
  if (typeof value === 'number') return excelDateFromSerial(value)
  if (typeof value === 'object' && value !== null) {
    if ('result' in value) return dateFromValue(value.result as ExcelJS.CellValue)
    if ('text' in value) return dateFromValue(value.text)
  }
  if (typeof value !== 'string') return null
  const parsed = new Date(value)
  return Number.isNaN(parsed.getTime()) ? null : parsed
}

function formatDate(value: ExcelJS.CellValue): string {
  const date = dateFromValue(value)
  if (!date) return ''
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`
}

function splitNombre(nombre: string): { apellidoPaterno: string; apellidoMaterno: string; nombres: string } {
  const partes = nombre.trim().split(/\s+/).filter(Boolean)
  return {
    apellidoPaterno: partes[0] ?? '',
    apellidoMaterno: partes[1] ?? '',
    nombres: partes.slice(2).join(' ') || nombre.trim(),
  }
}

type SheetExtent = { lastRow: number; columns: number }

const SHEET_EXTENT_CACHE = new WeakMap<ExcelJS.Worksheet, SheetExtent>()

function sheetExtent(ws: ExcelJS.Worksheet): SheetExtent {
  const cached = SHEET_EXTENT_CACHE.get(ws)
  if (cached) return cached

  let lastRow = 0
  let columns = 0
  ws.eachRow({ includeEmpty: false }, (row) => {
    if (row.number > lastRow) lastRow = row.number
    if (row.cellCount > columns) columns = row.cellCount
  })

  const extent: SheetExtent = {
    lastRow: lastRow > 0 ? lastRow : (ws.rowCount || 0),
    columns: Math.min(Math.max(columns, 1), 1000),
  }
  SHEET_EXTENT_CACHE.set(ws, extent)
  return extent
}

function maxColumn(ws: ExcelJS.Worksheet): number {
  return sheetExtent(ws).columns
}

function buildHeaderPath(ws: ExcelJS.Worksheet, headerRow: number, col: number): string[] {
  const path: string[] = []
  for (let row = Math.max(1, headerRow - 3); row <= headerRow; row++) {
    const text = normalizeText(ws.getRow(row).getCell(col).value)
    if (text && path[path.length - 1] !== text) path.push(text)
  }
  return path
}

function looksLikeHeaderRow(ws: ExcelJS.Worksheet, row: number, catalogo: CatalogoConceptos): boolean {
  let signals = 0
  for (let col = 1; col <= maxColumn(ws); col++) {
    const text = normalizeText(ws.getRow(row).getCell(col).value)
    if (!text) continue
    const path = [text]
    if (documentScore(path) > 0) signals++
    if (isNameHeader(path, ['APELLIDO PATERNO', 'APELLIDO MATERNO', 'NOMBRES', 'NOMBRE', 'APELLIDOS Y NOMBRES'])) signals++
    if (isPeriodHeader(path) || isExplicitConceptHeader(path) || isAggregateHeader(path)) signals++
    if (classificationFor(text, catalogo).family !== 'desconocido') signals++
    if (/ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|SETIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE/.test(text)) signals++
  }
  return signals >= 2
}

function sheetDataLimit(ws: ExcelJS.Worksheet): number {
  return sheetExtent(ws).lastRow
}

function findCandidates(workbook: ExcelJS.Workbook, catalogo: CatalogoConceptos): CandidateTable[] {
  const candidates = new Map<string, CandidateTable>()

  for (const ws of workbook.worksheets) {
    const dataLimit = sheetDataLimit(ws)
    const maxRow = Math.min(dataLimit, 60)
    const columns = maxColumn(ws)

    for (let headerRow = 1; headerRow <= maxRow; headerRow++) {
      if (!looksLikeHeaderRow(ws, headerRow, catalogo)) continue
      let dniCol: number | null = null
      let dniScore = 0
      for (let col = 1; col <= columns; col++) {
        const score = documentScore(buildHeaderPath(ws, headerRow, col))
        if (score > dniScore) {
          dniScore = score
          dniCol = col
        }
      }
      if (dniCol === null || dniScore <= 0) continue

      let dataRow = 0
      const lastProbe = Math.min(dataLimit, headerRow + 200)
      for (let row = headerRow + 1; row <= lastProbe; row++) {
        if (normalizeDni(ws.getRow(row).getCell(dniCol).value)) {
          dataRow = row
          break
        }
      }
      if (!dataRow) continue

      const key = `${ws.name}|${dniCol}|${dataRow}`
      const current = candidates.get(key)
      if (current) {
        current.headerRows.push(headerRow)
        for (let col = 1; col <= columns; col++) {
          const path = buildHeaderPath(ws, headerRow, col)
          const existing = current.paths.get(col) ?? []
          if (path.length > existing.length) current.paths.set(col, path)
        }
        continue
      }

      const paths = new Map<number, string[]>()
      for (let col = 1; col <= columns; col++) {
        const path = buildHeaderPath(ws, headerRow, col)
        if (path.length > 0) paths.set(col, path)
      }
      candidates.set(key, { ws, headerRows: [headerRow], dataRow, dniCol, paths })
    }
  }

  return [...candidates.values()]
}

function buildColumnInfos(table: CandidateTable, catalogo: CatalogoConceptos): ColumnInfo[] {
  const columns: ColumnInfo[] = []
  for (const [col, path] of table.paths) {
    const classification = classificationFromPath(path, catalogo)
    const name = nameKind(path)
    columns.push({
      col,
      path,
      label: path[path.length - 1] ?? '',
      identity: col === table.dniCol || documentScore(path) > 0,
      nameApellido: name === 'apellido',
      nameMaterno: name === 'materno',
      nameNombres: name === 'nombres',
      nameCompleto: name === 'completo',
      period: isPeriodHeader(path),
      explicitConcept: isExplicitConceptHeader(path),
      knownConcept: classification.family !== 'desconocido' && classification.family !== 'agregado',
      aggregate: isAggregateHeader(path),
      amount: isAmountHeader(path, classification),
      month: periodFromHeader(path),
      measure: measureFromText(path.join(' ')),
    })
  }
  return columns
}

function pareceTablaDeDias(table: CandidateTable, columns: ColumnInfo[]): boolean {
  const monthColumns = columns.filter((column) => column.month && !column.period)
  if (monthColumns.length === 0) return false
  if (columns.some((column) => column.amount || column.explicitConcept || column.knownConcept)) return false

  let filas = 0
  let dias = 0
  let dinero = 0
  const hasta = Math.min(sheetDataLimit(table.ws), table.dataRow + 40)
  for (let row = table.dataRow; row <= hasta; row++) {
    if (!normalizeDni(table.ws.getRow(row).getCell(table.dniCol).value)) continue
    filas++
    for (const column of monthColumns) {
      const value = toNumberNullable(table.ws.getRow(row).getCell(column.col).value)
      if (value === null) continue
      if (Number.isInteger(value) && value >= 0 && value <= 31) dias++
      else if (value !== 0) dinero++
    }
    if (filas >= 30) break
  }

  return filas > 0 && dias >= dinero
}

type WorkerMeta = {
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string
  fechaInicio: string
  fechaCese: string
}

type ScanResult = {
  registros: RegistroRemuneracion[]
  periodos: Map<string, Set<string>>
  meta: Map<string, WorkerMeta>
}

function metaFromRow(columns: ColumnInfo[], row: ExcelJS.Row): WorkerMeta {
  const apellido = columns.find((column) => column.nameApellido)
  const materno = columns.find((column) => column.nameMaterno)
  const nombres = columns.find((column) => column.nameNombres)
  const completo = columns.find((column) => column.nameCompleto)
  const ingreso = columns.find((column) => dateColumn(column.path, 'ingreso'))
  const cese = columns.find((column) => dateColumn(column.path, 'cese'))

  let nombreCompleto = completo ? normalizeText(row.getCell(completo.col).value) : ''
  if (!nombreCompleto && (apellido || materno || nombres)) {
    nombreCompleto = [
      apellido ? normalizeText(row.getCell(apellido.col).value) : '',
      materno ? normalizeText(row.getCell(materno.col).value) : '',
      nombres ? normalizeText(row.getCell(nombres.col).value) : '',
    ].filter(Boolean).join(' ')
  }

  const split = nombreCompleto ? splitNombre(nombreCompleto) : { apellidoPaterno: '', apellidoMaterno: '', nombres: '' }
  return {
    apellidoPaterno: apellido ? normalizeText(row.getCell(apellido.col).value) : split.apellidoPaterno,
    apellidoMaterno: materno ? normalizeText(row.getCell(materno.col).value) : split.apellidoMaterno,
    nombres: nombres ? normalizeText(row.getCell(nombres.col).value) : split.nombres,
    fechaInicio: ingreso ? formatDate(row.getCell(ingreso.col).value) : '',
    fechaCese: cese ? formatDate(row.getCell(cese.col).value) : '',
  }
}

function provenanceFor(
  loteId: string,
  ws: ExcelJS.Worksheet,
  row: number,
  column: number,
  path: string[],
  value: ExcelJS.CellValue,
): Provenance {
  return {
    loteId,
    hoja: ws.name,
    fila: row,
    columna: column,
    rutaEncabezado: path,
    valorOriginal: normalizeText(value),
  }
}

function recordId(
  loteId: string,
  ws: ExcelJS.Worksheet,
  row: number,
  column: number,
  dni: string,
  periodo: PeriodoAtomico | null,
  conceptKey: string,
  medida: string,
): string {
  return [loteId, ws.name, row, column, dni, periodo?.key ?? 'sin-periodo', conceptKey, medida].join('|')
}

function scanTable(table: CandidateTable, columns: ColumnInfo[], options: OpcionesRemuneracionAtomica, loteId: string): ScanResult {
  const catalogo = options.catalogo ?? CATALOGO_CONCEPTOS_POR_DEFECTO
  const registros: RegistroRemuneracion[] = []
  const periodos = new Map<string, Set<string>>()
  const meta = new Map<string, WorkerMeta>()

  if (pareceTablaDeDias(table, columns)) return { registros, periodos, meta }

  const dniColumn = columns.find((column) => column.col === table.dniCol)
  if (!dniColumn) return { registros, periodos, meta }

  const periodColumn = columns.find((column) => column.period && !column.identity)
  const explicitConceptColumn = columns.find((column) => column.explicitConcept)
  const knownConceptColumns = columns.filter((column) => column.knownConcept && !column.explicitConcept)
  const measureColumn = columns.find((column) => /MEDIDA/.test(normalizeText(column.path.join(' '))))
  const monthColumns = columns.filter((column) => column.month && !column.period)
  let monthColumnHasNumbers = false
  const dataLimit = sheetDataLimit(table.ws)
  for (const column of monthColumns) {
    for (let row = table.dataRow; row < Math.min(dataLimit, table.dataRow + 30); row++) {
      if (normalizeDni(table.ws.getRow(row).getCell(table.dniCol).value) && toNumberNullable(table.ws.getRow(row).getCell(column.col).value) !== null) {
        monthColumnHasNumbers = true
        break
      }
    }
    if (monthColumnHasNumbers) break
  }
  const hasTotalSource = columns.some((column) => column.aggregate && column.amount)
  const hasDetailSource = Boolean(periodColumn || explicitConceptColumn || (monthColumns.length > 0 && monthColumnHasNumbers) || knownConceptColumns.length === 1 || hasTotalSource)
  if (!hasDetailSource) return { registros, periodos, meta }

  const amountColumns = columns.filter((column) =>
    column.col !== table.dniCol &&
    !column.nameApellido &&
    !column.nameMaterno &&
    !column.nameNombres &&
    !column.nameCompleto &&
    !column.period &&
    (column.amount || column.month)
  )

  if (amountColumns.length === 0) return { registros, periodos, meta }

  const clasificacionPorColumna = new Map<number, ConceptClassification>()
  for (const column of columns) clasificacionPorColumna.set(column.col, classificationFromPath(column.path, catalogo))
  const clasificacionPorConcepto = new Map<string, ConceptClassification>()
  const clasificacionMensual: ConceptClassification = { key: 'remuneracion_mensual', label: 'Remuneración mensual', family: 'remunerativo' }

  const hasta = Math.min(dataLimit, table.dataRow + 5000)
  for (let rowNumber = table.dataRow; rowNumber <= hasta; rowNumber++) {
    const row = table.ws.getRow(rowNumber)
    const dni = normalizeDni(row.getCell(table.dniCol).value)
    if (!dni) continue

    const rowMeta = meta.has(dni) ? meta.get(dni)! : metaFromRow(columns, row)
    if (!meta.has(dni)) meta.set(dni, rowMeta)
    const rowPeriod = periodColumn ? periodFromValue(row.getCell(periodColumn.col).value) : null
    const rowConcept = explicitConceptColumn ? normalizeText(row.getCell(explicitConceptColumn.col).value) : ''

    for (const amountColumn of amountColumns) {
      const value = row.getCell(amountColumn.col).value
      const monto = toNumberNullable(value)
      const periodo = amountColumn.month ?? rowPeriod
      if (periodo?.mes) {
        const set = periodos.get(dni) ?? new Set<string>()
        set.add(periodo.mes)
        periodos.set(dni, set)
      }

      if (monto === null || monto === 0) continue

      const pathLabel = clasificacionPorColumna.get(amountColumn.col) ?? { key: 'desconocido', label: '', family: 'desconocido' as const }
      let classification: ConceptClassification
      if (amountColumn.month) {
        classification = clasificacionMensual
      } else if (explicitConceptColumn && amountColumn.col !== explicitConceptColumn.col) {
        const cached = clasificacionPorConcepto.get(rowConcept)
        classification = cached ?? classificationFor(rowConcept, catalogo)
        if (!cached) clasificacionPorConcepto.set(rowConcept, classification)
      } else {
        classification = pathLabel
      }
      const label = classification.label || rowConcept || amountColumn.label
      const aggregate = amountColumn.aggregate || classification.family === 'agregado' || /TOTAL|SUBTOTAL/.test(normalizeText(label))
      const origen = aggregate ? 'total' : amountColumn.month ? 'mensual' : 'concepto'
      const provenance = {
        ...provenanceFor(loteId, table.ws, rowNumber, amountColumn.col, amountColumn.path, value),
        origen,
        periodo: periodo?.label,
        concepto: label,
      }
      const id = recordId(loteId, table.ws, rowNumber, amountColumn.col, dni, periodo, classification.key, 'total')

      if (aggregate) {
        registros.push({
          id,
          loteId,
          dni,
          periodo,
          conceptKey: 'total_declarado',
          conceptLabel: label || 'Total declarado',
          medida: 'total',
          monto,
          moneda: 'PEN',
          rowKind: /SUBTOTAL/.test(normalizeText(label)) ? 'subtotal' : 'total',
          estado: 'considerado',
          origen: 'total',
          provenance,
        })
        continue
      }

      const medida = measureColumn
        ? measureFromText(row.getCell(measureColumn.col).value) ?? amountColumn.measure ?? 'devengado'
        : amountColumn.measure ?? 'devengado'
      const estado: EstadoRegistro = classification.family === 'desconocido'
        ? 'desconocido'
        : classification.family === 'no_remunerativo' || !['bruto', 'devengado'].includes(medida) || !currencyIsPen(`${amountColumn.path.join(' ')} ${rowConcept}`)
          ? 'excluido'
          : 'considerado'

      registros.push({
        id,
        loteId,
        dni,
        periodo,
        conceptKey: classification.key,
        conceptLabel: label,
        medida,
        monto,
        moneda: 'PEN',
        rowKind: 'detalle',
        estado,
        origen,
        provenance,
      })
    }
  }

  return { registros, periodos, meta }
}

function rankRegistro(registro: RegistroRemuneracion): number {
  if (registro.origen === 'concepto') return 3
  if (registro.origen === 'mensual') return 2
  return 1
}

function aggregateAtomic(
  scanned: ScanResult,
  options: OpcionesRemuneracionAtomica,
  loteId: string,
): { trabajadores: ProyeccionRemuneracion[]; registros: RegistroRemuneracion[] } {
  const registros = scanned.registros
  const facts = registros.filter((registro) => registro.estado === 'considerado' && registro.origen !== 'total')
  const declarados = registros.filter((registro) => registro.estado === 'considerado' && registro.origen === 'total')
  const byDni = new Map<string, RegistroRemuneracion[]>()

  for (const registro of facts) {
    const list = byDni.get(registro.dni) ?? []
    list.push(registro)
    byDni.set(registro.dni, list)
  }

  const selectedByDni = new Map<string, RegistroRemuneracion[]>()
  for (const [dni, list] of byDni) {
    const conceptPeriods = new Set(
      list.filter((registro) => registro.origen === 'concepto' && registro.periodo).map((registro) => registro.periodo!.key)
    )
    const active = list.filter((registro) => {
      if (registro.origen === 'mensual' && registro.periodo && conceptPeriods.has(registro.periodo.key)) {
        registro.estado = 'duplicado'
        return false
      }
      return true
    })
    const hasPeriod = active.some((registro) => registro.periodo)
    const withoutPeriod = active.filter((registro) => !registro.periodo)
    if (hasPeriod) {
      for (const registro of withoutPeriod) registro.estado = 'excluido'
    }

    const groups = new Map<string, RegistroRemuneracion[]>()
    for (const registro of active.filter((registro) => registro.estado === 'considerado')) {
      const key = `${registro.periodo?.key ?? 'sin-periodo'}|${registro.conceptKey}|${registro.medida}`
      const group = groups.get(key) ?? []
      group.push(registro)
      groups.set(key, group)
    }

    const selected: RegistroRemuneracion[] = []
    for (const group of groups.values()) {
      const ordered = [...group].sort((a, b) => rankRegistro(b) - rankRegistro(a) || a.provenance.fila - b.provenance.fila)
      selected.push(ordered[0])
      for (const duplicate of ordered.slice(1)) duplicate.estado = 'duplicado'
    }
    selectedByDni.set(dni, selected)
  }

  const declaredByDni = new Map<string, RegistroRemuneracion>()
  for (const registro of declarados) {
    const current = declaredByDni.get(registro.dni)
    if (!current) {
      declaredByDni.set(registro.dni, registro)
      continue
    }
    if (Math.abs(current.monto - registro.monto) > 0.01) registro.estado = 'duplicado'
    else registro.estado = 'duplicado'
  }

  const dnis = new Set<string>([
    ...byDni.keys(),
    ...declaredByDni.keys(),
    ...scanned.meta.keys(),
  ])
  const trabajadores: ProyeccionRemuneracion[] = []

  for (const dni of dnis) {
    const selected = selectedByDni.get(dni) ?? []
    const meta = scanned.meta.get(dni) ?? { apellidoPaterno: '', apellidoMaterno: '', nombres: '', fechaInicio: '', fechaCese: '' }
    const remuneraciones: MonthlyValues = { ...MESES_EN_CERO }
    let totalDesdeHechos = 0
    for (const registro of selected) {
      totalDesdeHechos += registro.monto
      if (registro.periodo?.mes && MONTHS.includes(registro.periodo.mes as (typeof MONTHS)[number])) {
        remuneraciones[registro.periodo.mes as keyof MonthlyValues] += registro.monto
      }
    }
    const available = scanned.periodos.get(dni) ?? new Set<string>()
    const mesesDisponibles = MONTHS.filter((mes) => available.has(mes))
    const mesesConValor = MONTHS.filter((mes) => remuneraciones[mes] !== 0)
    const declared = declaredByDni.get(dni)
    const totalDeclarado = declared?.monto
    const total = totalDesdeHechos !== 0 || selected.length > 0 ? totalDesdeHechos : (totalDeclarado ?? 0)
    const base: ResumenRemuneracion['base'] = selected.some((registro) => registro.periodo)
      ? (totalDeclarado !== undefined ? 'mixto' : 'detalle')
      : totalDeclarado !== undefined || selected.length > 0
        ? 'declarado'
        : 'declarado'
    const conceptos: Record<string, number> = {}
    for (const registro of selected) conceptos[registro.conceptKey] = (conceptos[registro.conceptKey] ?? 0) + registro.monto

    const considered = selected
    const excluded = registros.filter((registro) => registro.dni === dni && registro.estado === 'excluido')
    const unknown = registros.filter((registro) => registro.dni === dni && registro.estado === 'desconocido')
    const duplicates = registros.filter((registro) => registro.dni === dni && registro.estado === 'duplicado')
    const conflictos: string[] = []
    if (totalDeclarado !== undefined && totalDesdeHechos !== 0 && Math.abs(totalDeclarado - totalDesdeHechos) > 0.01) {
      conflictos.push(`Total declarado (${totalDeclarado}) difiere del detalle atómico (${totalDesdeHechos}).`)
    }

    const resumen: ResumenRemuneracion = {
      loteId,
      lotesAnteriores: options.lotesAnteriores ?? [],
      base,
      totalDesdeHechos,
      ...(totalDeclarado !== undefined ? { totalDeclarado } : {}),
      ...(totalDeclarado !== undefined ? { diferencia: totalDesdeHechos - totalDeclarado } : {}),
      mensualDisponible: mesesDisponibles.length > 0,
      considerados: considered.length,
      excluidos: excluded.length,
      duplicados: duplicates.length,
      conflictos,
      conceptosDesconocidos: [...new Set(unknown.map((registro) => registro.conceptLabel))],
      conceptosExcluidos: [...new Set(excluded.map((registro) => registro.conceptLabel))],
      conceptos,
      provenance: declared
        ? [...considered.map((registro) => registro.provenance), declared.provenance]
        : considered.map((registro) => registro.provenance),
    }

    const mesesPresentes = mesesConValor
    trabajadores.push({
      dni,
      ...meta,
      remuneraciones: { ...remuneraciones, total },
      remuneracionIncompleta: mesesDisponibles.length < MONTHS.length,
      mesesRemuneracionDisponibles: mesesDisponibles,
      mesesRemuneraciones: mesesPresentes,
      mesesPresentes,
      remuneracionResumen: resumen,
    })
  }

  return { trabajadores, registros }
}

export function detectarFormatoDetalle(
  workbook: ExcelJS.Workbook,
  catalogo: CatalogoConceptos = CATALOGO_CONCEPTOS_POR_DEFECTO,
): boolean {
  for (const table of findCandidates(workbook, catalogo)) {
    const columns = buildColumnInfos(table, catalogo)
    if (columns.some((column) => column.explicitConcept)) return true
    if (columns.some((column) => column.period && !column.identity)) return true
  }
  return false
}

export function extraerRemuneracionAtomico(
  workbook: ExcelJS.Workbook,
  options: OpcionesRemuneracionAtomica = {},
): ResultadoRemuneracionAtomica {
  const loteId = options.loteId ?? 'lote-actual'
  const catalogo = options.catalogo ?? CATALOGO_CONCEPTOS_POR_DEFECTO
  const scanned: ScanResult = { registros: [], periodos: new Map(), meta: new Map() }

  for (const table of findCandidates(workbook, catalogo)) {
    const columns = buildColumnInfos(table, catalogo)
    const result = scanTable(table, columns, { ...options, catalogo }, loteId)
    scanned.registros.push(...result.registros)
    for (const [dni, periods] of result.periodos) {
      const current = scanned.periodos.get(dni) ?? new Set<string>()
      for (const period of periods) current.add(period)
      scanned.periodos.set(dni, current)
    }
    for (const [dni, meta] of result.meta) {
      if (!scanned.meta.has(dni)) scanned.meta.set(dni, meta)
    }
  }

  const projected = aggregateAtomic(scanned, { ...options, catalogo }, loteId)
  const considered = projected.registros.filter((registro) => registro.estado === 'considerado')
  const modo: ResultadoRemuneracionAtomica['modo'] = considered.some((registro) => registro.origen === 'concepto' && registro.periodo)
    ? 'detalle'
    : considered.some((registro) => registro.origen === 'mensual' && registro.periodo)
      ? 'mensual'
      : considered.some((registro) => registro.origen === 'total' || registro.origen === 'concepto')
        ? 'declarado'
        : 'indisponible'
  return {
    disponible: projected.trabajadores.length > 0 && projected.registros.length > 0,
    modo,
    trabajadores: projected.trabajadores,
    registros: projected.registros,
  }
}

export { MONTHS as mesesAtomicos }

