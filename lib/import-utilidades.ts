import * as ExcelJS from 'exceljs'
import { detectarFormatoDetalle, extraerRemuneracionAtomico } from './import-utilidades-atomico'
import {
  clasificarTipo,
  extraerIncidencias,
  etiquetaMes,
  formatearPeriodo,
  leerPeriodo,
  type Confianza,
  type EstadoIncidencia,
  type OrigenCampo,
  type PeriodoIncidencia,
  type TipoIncidencia,
} from './import-incidencias'
import type { IncidenciaNormalizada } from './import-incidencias'
import type { OpcionesRemuneracionAtomica, RegistroRemuneracion, ResumenRemuneracion, ResultadoRemuneracionAtomica } from './import-utilidades-atomico'

export type { Confianza, EstadoIncidencia, OrigenCampo, PeriodoIncidencia, TipoIncidencia }
export { clasificarTipo, formatearPeriodo, leerPeriodo }

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

export type IncidenciaImportada = {
  dni: string
  codigo: string
  descripcion: string
  cantidadDias: number
  diasNeto?: number
  fechaInicio: ExcelJS.CellValue
  fechaFin: ExcelJS.CellValue
  periodo?: PeriodoIncidencia | null
  estado?: EstadoIncidencia
  motivo?: string
  tipo?: TipoIncidencia
  confianza?: Confianza
  origen?: OrigenCampo[]
}

export type ImportedWorker = {
  dni: string
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string
  fechaInicio: ExcelJS.CellValue
  fechaCese: ExcelJS.CellValue
  remuneraciones: MonthlyValues
  remuneracionIncompleta?: boolean
  mesesRemuneracionDisponibles?: string[]
  remuneracionResumen?: ResumenRemuneracion
  loteRemuneracion?: string
  diasTrabajados?: MonthlyValues
  mesesRemuneraciones?: string[]
  mesesDias?: string[]
  mesesPresentes?: string[]
  incidencias?: IncidenciaImportada[]
  diasNoLaborados?: number
  diasEfectivos?: number
  diasPosibles?: number
  diasNoLaboradosReferencia?: number
  diasEfectivosReferencia?: number
  jornada?: number
  feriados?: string[]
  laborablesPorMes?: MonthlyValues
}

export type MesesTrabajados = Pick<ImportedWorker, 'diasTrabajados' | 'mesesDias' | 'mesesPresentes'>

/**
 * Única fuente de verdad de los meses trabajados.
 *
 * Regla: un mes está trabajado si el período tiene días registrados > 0. Si el
 * archivo no trae desglose mensual de días, se recurre a la permanencia que el
 * propio motor ya consolidó (mesesDias y, en último caso, mesesPresentes).
 * Tener una fila en el Excel no implica haber trabajado ese mes.
 */
export function mesesTrabajados(worker: MesesTrabajados, override?: string[]): string[] {
  if (override) return months.filter((mes) => override.includes(mes))
  if (worker.diasTrabajados) return months.filter((mes) => (worker.diasTrabajados?.[mes] ?? 0) > 0)
  if (worker.mesesDias && worker.mesesDias.length > 0) return months.filter((mes) => worker.mesesDias?.includes(mes))
  return months.filter((mes) => (worker.mesesPresentes ?? []).includes(mes))
}

/** Primer valor con contenido: una cadena vacía es "sin dato", no un dato. */
function primerValor<T>(...candidatos: (T | null | undefined)[]): T {
  for (const candidato of candidatos) {
    if (candidato === null || candidato === undefined) continue
    if (typeof candidato === 'string' && candidato.trim() === '') continue
    return candidato
  }
  return '' as T
}

/**
 * Consolida varias cargas del mismo trabajador (por ejemplo, la carga de
 * remuneraciones y la de días) en un único registro.
 *
 * Reglas:
 *  1. Cada dato se toma de la primera fuente que lo aporta; una cadena vacía no
 *     es un dato, así que la fecha de ingreso de la planilla de remuneraciones
 *     no se pierde porque la carga de días la traiga vacía.
 *  2. Las remuneraciones se suman por fuente, nunca se derivan de los días.
 *  3. Las incidencias se concatenan: son la única fuente de días no laborados.
 *  4. Los meses trabajados se recalculan con mesesTrabajados, de modo que tener
 *     remuneración en un mes no lo convierte en mes trabajado.
 */
export function consolidarTrabajadores<T extends ImportedWorker>(...fuentes: T[][]): T[] {
  const porDni = new Map<string, T>()

  for (const fuente of fuentes) {
    for (const trabajador of fuente) {
      const dni = String(trabajador.dni).trim().padStart(8, '0')
      const existente = porDni.get(dni)

      if (!existente) {
        porDni.set(dni, { ...trabajador, dni } as T)
        continue
      }

      const remuneraciones =
        existente.remuneraciones.total > 0 ? existente.remuneraciones : trabajador.remuneraciones
      const diasTrabajados = existente.diasTrabajados ?? trabajador.diasTrabajados
      const mesesDias = months.filter(
        (mes) => (existente.mesesDias ?? []).includes(mes) || (trabajador.mesesDias ?? []).includes(mes)
      )
      const mesesPresentes = months.filter(
        (mes) =>
          (existente.mesesPresentes ?? []).includes(mes) ||
          (trabajador.mesesPresentes ?? []).includes(mes)
      )

      porDni.set(dni, {
        ...existente,
        ...trabajador,
        dni,
        apellidoPaterno: primerValor(existente.apellidoPaterno, trabajador.apellidoPaterno),
        apellidoMaterno: primerValor(existente.apellidoMaterno, trabajador.apellidoMaterno),
        nombres: primerValor(existente.nombres, trabajador.nombres),
        fechaInicio: primerValor(existente.fechaInicio, trabajador.fechaInicio),
        fechaCese: primerValor(existente.fechaCese, trabajador.fechaCese),
        remuneraciones,
        diasTrabajados,
        incidencias: [...(existente.incidencias ?? []), ...(trabajador.incidencias ?? [])],
        mesesRemuneraciones: months.filter(
          (mes) =>
            (existente.mesesRemuneraciones ?? []).includes(mes) ||
            (trabajador.mesesRemuneraciones ?? []).includes(mes)
        ),
        mesesDias,
        mesesPresentes: mesesTrabajados({ diasTrabajados, mesesDias, mesesPresentes }),
      } as T)
    }
  }

  return [...porDni.values()]
}

const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'] as const

const MESES_EN_CERO: MonthlyValues = {
  enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0,
  julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0,
  total: 0,
}

const normalizeHeader = (value: ExcelJS.CellValue): string => {
  const text =
    typeof value === 'object' && value !== null && 'text' in value
      ? String(value.text ?? '')
      : String(value ?? '')

  const normalized = text
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase()

  return normalized === 'SETIEMBRE' ? 'SEPTIEMBRE' : normalized
}

const normalizeDni = (value: ExcelJS.CellValue): string | null => {
  const dni = String(value ?? '').trim().replace(/\.0$/, '')

  if (/^\d{9}$/.test(dni)) return dni

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

const parseIncFecha = (value: ExcelJS.CellValue): Date | null => celdaValorFecha(value)

export { parseIncFecha, months, MESES_EN_CERO }

// ---------------------------------------------------------------------------
// Detección genérica de tablas por encabezados
// ---------------------------------------------------------------------------

type TipoCampo =
  | 'dni'
  | 'apellidoPaterno'
  | 'apellidoMaterno'
  | 'nombres'
  | 'nombreCompleto'
  | 'codigo'
  | 'jornada'
  | 'fechaIngreso'
  | 'fechaCese'
  | 'descripcion'
  | 'cantidad'
  | 'fechaI'
  | 'fechaF'
  | 'diasPosibles'
  | 'diasDeduccion'
  | 'diasEfectivos'
  | 'diasNeto'
  | 'remuneracionTotal'

const SINONIMOS: Record<TipoCampo, (string | RegExp)[]> = {
  dni: ['DNI', 'NUMERO', 'NRO', 'NRO.', 'NUMERO DE IDENTIDAD', 'DOCUMENTO DE IDENTIDAD', 'NUMERO DE DOCUMENTO', 'NRO. DOC. IDENTIDAD', 'NUMDOCIDE'],
  apellidoPaterno: ['APELLIDO PATERNO'],
  apellidoMaterno: ['APELLIDO MATERNO'],
  nombres: ['NOMBRES', 'NOMBRE'],
  nombreCompleto: ['APELLIDOS Y NOMBRES', 'NOMBRE COMPLETO', 'NOMBRES Y APELLIDOS'],
  codigo: ['CODIGO', 'COD.'],
  jornada: ['JORNADA', 'JORNADA LABORAL'],
  fechaIngreso: ['FECHA DE INGRESO', 'FECHA INGRESO', 'FECHA DE INICIO', 'FECHA INICIO', 'F. INGRESO', 'F. INICIO', 'FECHA INICIAL', 'F.INICIO'],
  fechaCese: ['FECHA DE CESE', 'FECHA CESE', 'F. CESE', 'F.FIN', 'F. FIN', 'FECHA FINAL', 'FECHA FIN', 'FECHA DE FIN'],
  descripcion: ['DESCRIPCION', 'MOTIVO', 'CONCEPTO', 'CONCEPTO DE FALTA', 'DETALLE', 'RAZON'],
  cantidad: ['NUM. DIAS', 'NUM DIAS', 'NRO DIAS', 'NUMERO DE DIAS', 'CANTIDAD', 'CANTIDAD DE DIAS', 'N. DIAS', 'N DIAS', 'NO DE DIAS'],
  fechaI: ['F. INICIO', 'F.INICIO', 'FECHA INICIAL', 'FECHA DESDE', 'DESDE'],
  fechaF: ['F. FIN', 'F.FIN', 'FECHA FINAL', 'FECHA HASTA', 'HASTA'],
  diasPosibles: [/DIAS FULL/, 'DIAS POSIBLES', 'DIAS LABORABLES', 'TOTAL DIAS DEL ANIO'],
  diasDeduccion: [/DEDUCCI/],
  diasEfectivos: [/^DIAS\s+\d{4}$/, 'DIAS TRABAJADOS'],
  diasNeto: ['DIAS EFECTIVOS TOTAL', 'DIAS NETOS', 'DIAS NETO', 'NETO'],
  remuneracionTotal: ['REMUNERACION', 'REMUNERACIONES', 'TOTAL REMUNERACION', 'TOTAL REMUNERACIONES', 'TOTAL INGRESOS', 'TOTAL DEVENGADO', 'REMUN. BASICA', 'SUELDO'],
}

const compactHeader = (value: string): string => value.replace(/[^A-Z0-9]/g, '')

function puntajeDocumento(texto: string): number {
  if (/^(TIPO|TYPE)\b/.test(texto)) return 0

  const compacto = compactHeader(texto)
  if (compacto === 'DNI') return 100
  if (['NRODOCIDE', 'NRODOCIDENTIDAD', 'NUMDOCIDE', 'NUMERODEDOCUMENTO', 'DOCUMENTODEIDENTIDAD', 'NUMERODEIDENTIDAD', 'NDOCUMENTO', 'NRODOCUMENTO'].includes(compacto)) return 90
  if (compacto === 'NUMERO' || compacto === 'NRO') return 30

  const tieneDocumento = /DOCUMENTO|DOCUMENT|DOC/.test(texto)
  const tieneNumero = /(^|\s)(NRO|NUMERO|NUM|N)(\s|$)/.test(texto) || compacto.startsWith('NRO') || compacto.startsWith('NUM')
  return tieneDocumento && (tieneNumero || /IDENT/.test(texto)) ? 80 : 0
}

function esEncabezadoDocumento(texto: string): boolean {
  return puntajeDocumento(texto) > 0
}

function puntajeRemuneracion(texto: string): number {
  if (!texto || /PERIODO|CRITERIO|CONCEPTO DE FALTA/.test(texto)) return 0
  if (/DOCUMENTO INGRESO|INGRESO CANTIDAD|INGRESOS Y SALIDAS|CANTIDAD.*INGRESO|FECHA.*INGRES|INGRES.*FECHA/.test(texto)) return 0

  const esTotal = /(^|\s)TOTAL(\s|$)/.test(texto)
  const esRemuneracion = /(^|\s)REM/.test(texto)
  const esSueldo = /(^|\s)SUELDO(\s|$)/.test(texto)
  const esDevengado = /DEVENGAD/.test(texto)
  const esIngreso = /INGRES/.test(texto)
  if (!esRemuneracion && !esSueldo && !esDevengado && !esIngreso) return 0
  if (esTotal) return 100
  if (esRemuneracion && /BASIC/.test(texto)) return 90
  if (esRemuneracion) return 80
  if (esSueldo) return 70
  if (esDevengado) return 65
  return 60
}

function coincideEncabezado(tipo: TipoCampo, texto: string): boolean {
  if (tipo === 'dni') return esEncabezadoDocumento(texto)
  if (tipo === 'remuneracionTotal') return puntajeRemuneracion(texto) > 0
  return SINONIMOS[tipo].some((x) => (typeof x === 'string' ? x === texto : x.test(texto)))
}

type EncabezadoAnalizado = {
  fila: number
  campos: Map<TipoCampo, number>
  semanas: Map<string, number>
  totalCol: number | null
  feriadoCol: number | null
}

function analizarFila(ws: ExcelJS.Worksheet, fila: number): EncabezadoAnalizado | null {
  const row = ws.getRow(fila)
  const agrupados = new Map<string, number[]>()
  row.eachCell((cell, col) => {
    const texto = normalizeHeader(cell.value)
    if (!texto) return
    const lista = agrupados.get(texto) ?? []
    lista.push(col)
    agrupados.set(texto, lista)
  })
  if (agrupados.size === 0) return null

  const campos = new Map<TipoCampo, number>()
  for (const tipo of Object.keys(SINONIMOS) as TipoCampo[]) {
    if (tipo === 'dni' || tipo === 'remuneracionTotal') {
      let mejor: { col: number; puntaje: number } | null = null
      for (const [texto, cols] of agrupados) {
        const puntaje = tipo === 'dni' ? puntajeDocumento(texto) : puntajeRemuneracion(texto)
        if (puntaje <= 0) continue
        for (const col of cols) {
          if (!mejor || puntaje > mejor.puntaje || (puntaje === mejor.puntaje && col < mejor.col)) {
            mejor = { col, puntaje }
          }
        }
      }
      if (mejor) campos.set(tipo, mejor.col)
      continue
    }

    let mejorCol: number | null = null
    for (const [texto, cols] of agrupados) {
      if (!coincideEncabezado(tipo, texto)) continue
      for (const col of cols) {
        if (mejorCol === null || col < mejorCol) mejorCol = col
      }
    }
    if (mejorCol !== null) campos.set(tipo, mejorCol)
  }

  const mesesEnFila: { indice: number; col: number }[] = []
  for (const [texto, cols] of agrupados) {
    const indice = months.findIndex((m) => m.toUpperCase() === texto)
    if (indice >= 0) {
      for (const col of cols) mesesEnFila.push({ indice, col })
    }
  }
  mesesEnFila.sort((a, b) => a.col - b.col)

  const semanas = new Map<string, number>()
  if (mesesEnFila.length > 0) {
    let mejorRun: { indice: number; col: number }[] = []
    let actualRun: { indice: number; col: number }[] = []
    for (const m of mesesEnFila) {
      const prev = actualRun[actualRun.length - 1]
      if (
        prev &&
        m.col === prev.col + 1 &&
        (m.indice === prev.indice + 1 || (m.indice === 0 && prev.indice === 11))
      ) {
        actualRun.push(m)
      } else {
        if (actualRun.length >= mejorRun.length) mejorRun = actualRun
        actualRun = [m]
      }
    }
    if (actualRun.length >= mejorRun.length) mejorRun = actualRun
    for (const m of mejorRun) semanas.set(months[m.indice], m.col)
  }

  let totalCol: number | null = null
  for (const [texto, cols] of agrupados) {
    if (!texto.includes('TOTAL')) continue
    for (const col of cols) {
      totalCol = totalCol === null ? col : Math.max(totalCol, col)
    }
  }

  let feriadoCol: number | null = null
  for (const [texto, cols] of agrupados) {
    if (!texto.includes('FERIADO')) continue
    for (const col of cols) {
      feriadoCol = feriadoCol === null ? col : Math.min(feriadoCol, col)
    }
  }

  return { fila, campos, semanas, totalCol, feriadoCol }
}

type Tabla = {
  ws: ExcelJS.Worksheet
  filaEncabezado: number
  filaDatos: number
  campos: Map<TipoCampo, number>
  semanas: Map<string, number>
  totalCol: number | null
  feriadoCol: number | null
  tipo: 'mensual' | 'remuneraciones' | 'incidencias' | 'consolidado'
}

function filaTieneDniValido(ws: ExcelJS.Worksheet, fila: number, hastaCol: number): boolean {
  const row = ws.getRow(fila)
  let valido = false
  row.eachCell((cell, col) => {
    if (col <= hastaCol && normalizeDni(cell.value)) valido = true
  })
  return valido
}

function detectarTablas(workbook: ExcelJS.Workbook): Tabla[] {
  const candidatos: (EncabezadoAnalizado & { ws: ExcelJS.Worksheet; key: string })[] = []

  for (const ws of workbook.worksheets) {
    const limiteEncabezados = Math.min(ws.rowCount, 60)
    for (let r = 1; r <= limiteEncabezados; r++) {
      const enc = analizarFila(ws, r)
      if (!enc || !enc.campos.has('dni')) continue

      const esMensual = enc.semanas.size >= 1
      const tieneIncidencia = enc.campos.has('descripcion') || enc.campos.has('cantidad')
      const tieneRemuneracion = enc.campos.has('remuneracionTotal')
      const tieneConsolidado =
        enc.campos.has('diasPosibles') ||
        enc.campos.has('jornada') ||
        enc.campos.has('diasEfectivos') ||
        enc.campos.has('diasDeduccion')

      if (!(esMensual || tieneIncidencia || tieneConsolidado || tieneRemuneracion)) continue

      const hastaProbe = Math.min(ws.rowCount, r + 500)
      let filaDatos = -1
      for (let dr = r + 1; dr <= hastaProbe; dr++) {
        if (filaTieneDniValido(ws, dr, 40)) {
          filaDatos = dr
          break
        }
      }
      if (filaDatos < 0) continue

      candidatos.push({ ...enc, ws, fila: r, key: `${ws.name}#${filaDatos}` })
    }
  }

  type CandidatoTabla = EncabezadoAnalizado & { ws: ExcelJS.Worksheet; key?: string }
  const porLlave = new Map<string, CandidatoTabla[]>()
  for (const c of candidatos) {
    const grp = porLlave.get(c.key) ?? []
    grp.push({ ...c })
    porLlave.set(c.key, grp)
  }

  const tablas: Tabla[] = []
  for (const key of porLlave.keys()) {
    const grp = porLlave.get(key)!
    const filaDatos = Number(key.split('#')[1])
    const ws = grp[0].ws

    const campos = new Map<TipoCampo, number>()
    const semanas = new Map<string, number>()
    let totalCol: number | null = null
    let feriadoCol: number | null = null

    for (const candidato of [...grp].sort((a, b) => b.fila - a.fila)) {
      for (const [tipo, col] of candidato.campos) {
        if (!campos.has(tipo) && col !== undefined) campos.set(tipo, col)
      }
      if (candidato.semanas.size > 0 && semanas.size === 0) {
        for (const [m, col] of candidato.semanas) semanas.set(m, col)
      }
      if (totalCol === null) totalCol = candidato.totalCol
      if (feriadoCol === null) feriadoCol = candidato.feriadoCol
    }

    const esMensual = semanas.size >= 1
    const tieneIncidencia = campos.has('descripcion') || campos.has('cantidad')
    const tieneRemuneracion = campos.has('remuneracionTotal')
    let tipo: Tabla['tipo'] = 'consolidado'
    if (esMensual) tipo = 'mensual'
    else if (tieneIncidencia) tipo = 'incidencias'
    else if (tieneRemuneracion) tipo = 'remuneraciones'

    const filaEncabezado = Math.min(...grp.map((g) => g.fila))

    if (tipo === 'incidencias') {
      let cFechaIOverride: number | null = null
      let cFechaFOverride: number | null = null
      for (let hr = filaEncabezado; hr < filaDatos; hr++) {
        const row = ws.getRow(hr)
        row.eachCell((cell, col) => {
          const texto = normalizeHeader(cell.value)
          if (texto && texto.includes('F. INICIO')) {
            cFechaIOverride = Math.min(cFechaIOverride ?? Infinity, col)
          }
          if (texto && texto.includes('F. FIN')) {
            cFechaFOverride = Math.min(cFechaFOverride ?? Infinity, col)
          }
        })
      }
      if (cFechaIOverride !== null) campos.set('fechaI', cFechaIOverride)
      if (cFechaFOverride !== null) campos.set('fechaF', cFechaFOverride)
    }

    tablas.push({ ws, filaEncabezado, filaDatos, campos, semanas, totalCol, feriadoCol, tipo })
  }

  return tablas
}

// ---------------------------------------------------------------------------
// Utilidades compartidas
// ---------------------------------------------------------------------------

function clasificarMensual(tabla: Tabla): 'remuneraciones' | 'diasMensuales' {
  const cols = [...tabla.semanas.values()]
  let dinero = 0
  let dias = 0
  let filasLeidas = 0

  for (let dr = tabla.filaDatos; dr < tabla.filaDatos + 500 && dr <= tabla.ws.rowCount; dr++) {
    const row = tabla.ws.getRow(dr)
    if (!normalizeDni(row.getCell(tabla.campos.get('dni')!).value)) continue
    filasLeidas++
    for (const col of cols) {
      const v = toNumber(row.getCell(col).value)
      // Un cero no es evidencia de nada: un mes sin dato aparece en cero tanto
      // en el cuadro de días como en el de remuneración, así que no se cuenta
      // para decidir el tipo de la matriz.
      if (v === 0) continue
      if (Number.isInteger(v) && v >= 0 && v <= 31) dias++
      else dinero++
    }
    if (filasLeidas >= 30) break
  }

  if (filasLeidas === 0) return 'remuneraciones'
  if (dias === 0 && dinero === 0) return 'diasMensuales'
  return dias >= dinero ? 'diasMensuales' : 'remuneraciones'
}

function calcularTotalSuma(tabla: Tabla, row: ExcelJS.Row): number {
  return months.reduce((suma, mes) => {
    const col = tabla.semanas.get(mes)
    return suma + (col === undefined ? 0 : toNumber(row.getCell(col).value))
  }, 0)
}

type InfoNombres = { apellidoPaterno: string; apellidoMaterno: string; nombres: string }

function splitNombreCompleto(nombreCompleto: string): InfoNombres {
  const partes = nombreCompleto.split(/\s+/).filter(Boolean)
  return {
    apellidoPaterno: partes[0] ?? '',
    apellidoMaterno: partes[1] ?? '',
    nombres: partes.slice(2).join(' ') || nombreCompleto,
  }
}

function nombresDesdeTabla(tabla: Tabla, row: ExcelJS.Row): InfoNombres | null {
  const cAp = tabla.campos.get('apellidoPaterno')
  const cAm = tabla.campos.get('apellidoMaterno')
  const cNom = tabla.campos.get('nombres')
  const cFull = tabla.campos.get('nombreCompleto')

  if (cAp || cAm) {
    return {
      apellidoPaterno: cAp ? limpiarNombre(row.getCell(cAp).value) : '',
      apellidoMaterno: cAm ? limpiarNombre(row.getCell(cAm).value) : '',
      nombres: cNom ? limpiarNombre(row.getCell(cNom).value) : '',
    }
  }
  if (cNom) {
    return splitNombreCompleto(limpiarNombre(row.getCell(cNom).value))
  }
  if (cFull) {
    return splitNombreCompleto(limpiarNombre(row.getCell(cFull).value))
  }
  return null
}

function inferirAnio(tablas: Tabla[]): number {
  const conteo = new Map<number, number>()
  for (const tabla of tablas) {
    for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 100); dr++) {
      const row = tabla.ws.getRow(dr)
      row.eachCell((cell) => {
        const fecha = parseExcelDate(cell.value)
        if (fecha) {
          const anio = fecha.getUTCFullYear()
          conteo.set(anio, (conteo.get(anio) ?? 0) + 1)
        }
      })
      if (dr >= tabla.filaDatos + 10) break
    }
  }
  let mejor = 2025
  let max = 0
  for (const [anio, n] of conteo) {
    if (n > max) {
      max = n
      mejor = anio
    }
  }
  return mejor
}

/**
 * Año del ejercicio que el archivo describe, deducido de las fechas que
 * aparecen en él (el año más repetido). Se usa solo para etiquetar períodos que
 * el archivo entrega sin año (por ejemplo, una columna "FEBRERO"): si el
 * archivo no tiene ninguna fecha, se devuelve null y el período se muestra sin
 * año en lugar de inventar uno.
 */
export function inferirAnioEjercicio(workbook: ExcelJS.Workbook): number | null {
  const conteo = new Map<number, number>()
  for (const ws of workbook.worksheets) {
    const ultimas: number[] = []
    ws.eachRow({ includeEmpty: false }, (row) => {
      if (ultimas.length >= 60) return
      ultimas.push(row.number)
    })
    for (const numero of ultimas) {
      ws.getRow(numero).eachCell({ includeEmpty: false }, (cell) => {
        const fecha = parseExcelDate(cell.value)
        if (!fecha) return
        const anio = fecha.getUTCFullYear()
        if (anio < 1970 || anio > 2200) return
        conteo.set(anio, (conteo.get(anio) ?? 0) + 1)
      })
    }
  }

  let mejor: number | null = null
  let max = 0
  for (const [anio, n] of conteo) {
    if (n > max) {
      max = n
      mejor = anio
    }
  }
  return mejor
}

// ---------------------------------------------------------------------------
// EXTRAER REMUNERACIONES
// ---------------------------------------------------------------------------

export type ExtraccionRemuneraciones = {
  importados: ImportedWorker[]
  trabajadores: ImportedWorker[]
  loteId?: string
  modo?: 'atomico' | 'compatibilidad' | 'indisponible'
  registros?: RegistroRemuneracion[]
}

function extraerRemuneracionesClasica(
  workbook: ExcelJS.Workbook,
  actuales: ImportedWorker[],
  opciones: { loteId: string; lotesAnteriores: string[] },
): ExtraccionRemuneraciones {
  const tablas = detectarTablas(workbook)
  const tablasMensuales = tablas.filter((t) => t.tipo === 'mensual' && clasificarMensual(t) === 'remuneraciones')
  const tablasTotales = tablasMensuales.length === 0
    ? tablas.filter((t) => t.tipo === 'remuneraciones' || (t.semanas.size === 0 && t.campos.has('remuneracionTotal')))
    : []

  if (tablasMensuales.length === 0 && tablasTotales.length === 0) return { importados: [], trabajadores: actuales }

  const porDni = new Map<string, ImportedWorker>()

  const procesarTabla = (tabla: Tabla, soloFaltantes: boolean) => {
    const cDni = tabla.campos.get('dni')!
    const mesesDisponibles = months.filter((mes) => tabla.semanas.has(mes))
    const totalCol = tabla.campos.get('remuneracionTotal') ?? tabla.totalCol

    for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 5000); dr++) {
      const row = tabla.ws.getRow(dr)
      const dni = normalizeDni(row.getCell(cDni).value)
      if (!dni || (soloFaltantes && porDni.has(dni))) continue

      const nombres = nombresDesdeTabla(tabla, row)
      const remuneraciones = { ...MESES_EN_CERO }
      for (const mes of mesesDisponibles) {
        const col = tabla.semanas.get(mes)
        if (col !== undefined) remuneraciones[mes] = toNumber(row.getCell(col).value)
      }
      const totalLeido = totalCol !== null && totalCol !== undefined ? toNumber(row.getCell(totalCol).value) : 0
      remuneraciones.total = totalLeido || calcularTotalSuma(tabla, row)
      const tieneValorRemuneracion = totalLeido !== 0 || mesesDisponibles.some((mes) => remuneraciones[mes] !== 0)
      if (!tieneValorRemuneracion) continue

      const mesesRemuneraciones = months.filter((mes) => remuneraciones[mes] > 0)
      const hechosMensuales = mesesDisponibles
        .filter((mes) => remuneraciones[mes] !== 0)
        .map((mes) => ({ mes, monto: remuneraciones[mes], columna: tabla.semanas.get(mes) ?? 0 }))
      const totalDesdeHechos = hechosMensuales.reduce((acc, hecho) => acc + hecho.monto, 0)
      const hayDetalleMensual = hechosMensuales.length > 0
      const hayConflictoTotal = hayDetalleMensual && totalLeido !== 0 && Math.abs(totalDesdeHechos - totalLeido) > 0.01
      const resumen: ResumenRemuneracion = {
        loteId: opciones.loteId,
        lotesAnteriores: opciones.lotesAnteriores,
        base: hayDetalleMensual ? 'detalle' : 'declarado',
        totalDesdeHechos,
        ...(totalLeido ? { totalDeclarado: totalLeido } : {}),
        ...(hayDetalleMensual && totalLeido ? { diferencia: totalDesdeHechos - totalLeido } : {}),
        mensualDisponible: mesesDisponibles.length > 0,
        considerados: hechosMensuales.length,
        excluidos: 0,
        duplicados: 0,
        conflictos: hayConflictoTotal
          ? [`Total declarado (${totalLeido}) difiere del detalle mensual (${totalDesdeHechos}).`]
          : [],
        conceptosDesconocidos: [],
        conceptosExcluidos: [],
        conceptos: hechosMensuales.length > 0 ? { remuneracion_mensual: totalDesdeHechos } : {},
        provenance: [
          ...hechosMensuales.map((hecho) => ({
            hoja: tabla.ws.name,
            fila: dr,
            columna: hecho.columna,
            rutaEncabezado: [tabla.ws.name, String(hecho.columna)],
            valorOriginal: String(row.getCell(hecho.columna).value ?? ''),
            periodo: hecho.mes,
            origen: 'mensual',
            concepto: 'Remuneración mensual',
            loteId: opciones.loteId,
          })),
          ...(hechosMensuales.length === 0 && totalLeido && totalCol !== null && totalCol !== undefined
            ? [{
                hoja: tabla.ws.name,
                fila: dr,
                columna: totalCol,
                rutaEncabezado: [tabla.ws.name, String(totalCol)],
                valorOriginal: String(row.getCell(totalCol).value ?? ''),
                origen: 'total',
                concepto: 'Total declarado',
                loteId: opciones.loteId,
              }]
            : []),
        ],
      }
      const nuevo: ImportedWorker = {
        dni,
        apellidoPaterno: nombres?.apellidoPaterno ?? '',
        apellidoMaterno: nombres?.apellidoMaterno ?? '',
        nombres: nombres?.nombres ?? '',
        fechaInicio: tabla.campos.get('fechaIngreso') ? formatearFecha(celdaValorFecha(row.getCell(tabla.campos.get('fechaIngreso')!).value)) : '',
        fechaCese: tabla.campos.get('fechaCese') ? formatearFecha(celdaValorFecha(row.getCell(tabla.campos.get('fechaCese')!).value)) : '',
        remuneraciones,
        remuneracionIncompleta: mesesDisponibles.length < months.length,
        mesesRemuneracionDisponibles: mesesDisponibles,
        mesesRemuneraciones,
        mesesPresentes: mesesRemuneraciones,
        remuneracionResumen: resumen,
        loteRemuneracion: opciones.loteId,
      }

      const existente = porDni.get(dni)
      if (!existente) {
        porDni.set(dni, nuevo)
        continue
      }

      for (const mes of months) {
        existente.remuneraciones[mes] = existente.remuneraciones[mes] + remuneraciones[mes]
      }
      existente.remuneraciones.total = existente.remuneraciones.total + remuneraciones.total
      const disponibles = new Set([
        ...(existente.mesesRemuneracionDisponibles ?? []),
        ...mesesDisponibles,
      ])
      existente.mesesRemuneracionDisponibles = months.filter((mes) => disponibles.has(mes))
      existente.remuneracionIncompleta = existente.mesesRemuneracionDisponibles.length < months.length
      existente.mesesRemuneraciones = months.filter((mes) => existente.remuneraciones[mes] > 0)
      existente.mesesPresentes = existente.mesesRemuneraciones
      const previo = existente.remuneracionResumen
      if (previo) {
        const conceptos: Record<string, number> = { ...previo.conceptos }
        for (const [concepto, monto] of Object.entries(resumen.conceptos)) {
          conceptos[concepto] = (conceptos[concepto] ?? 0) + monto
        }
        existente.remuneracionResumen = {
          ...previo,
          totalDesdeHechos: previo.totalDesdeHechos + resumen.totalDesdeHechos,
          mensualDisponible: previo.mensualDisponible || resumen.mensualDisponible,
          considerados: previo.considerados + resumen.considerados,
          duplicados: previo.duplicados + 1,
          conflictos: [...previo.conflictos, ...resumen.conflictos],
          conceptos,
          provenance: [...previo.provenance, ...resumen.provenance],
        }
        if (previo.totalDeclarado !== undefined) {
          existente.remuneracionResumen.diferencia = existente.remuneracionResumen.totalDesdeHechos - previo.totalDeclarado
        }
      } else {
        existente.remuneracionResumen = resumen
      }
    }
  }

  for (const tabla of tablasMensuales) procesarTabla(tabla, false)
  for (const tabla of tablasTotales) procesarTabla(tabla, true)

  const trabajadoresImportados = [...porDni.values()]

  const trabajadores = trabajadoresImportados.map((rem) => {
    const existente = actuales.find(
      (trabajador) => String(trabajador.dni).trim().padStart(8, '0') === rem.dni
    )

    if (!existente) return rem

    return {
      ...existente,
      ...rem,
      diasTrabajados: existente.diasTrabajados,
      mesesDias: existente.mesesDias,
      mesesPresentes: Array.from(
        new Set([
          ...(rem.mesesRemuneraciones ?? []),
          ...(existente.mesesDias ?? []),
        ])
      ),
    }
  })

  return { importados: trabajadoresImportados, trabajadores }
}

export function extraerRemuneraciones(
  workbook: ExcelJS.Workbook,
  actuales: ImportedWorker[],
  opciones: OpcionesRemuneracionAtomica = {},
): ExtraccionRemuneraciones {
  const loteId = opciones.loteId ?? 'lote-actual'
  const lotesAnteriores = [
    ...new Set(
      actuales
        .map((trabajador) => trabajador.loteRemuneracion)
        .filter((lote): lote is string => Boolean(lote) && lote !== loteId)
    ),
  ]
  const clasico = extraerRemuneracionesClasica(workbook, actuales, { loteId, lotesAnteriores })
  const hayDetalle = detectarFormatoDetalle(workbook, opciones.catalogo)
  const requiereAtomico = hayDetalle || clasico.importados.length === 0
  const atomico: ResultadoRemuneracionAtomica = requiereAtomico
    ? extraerRemuneracionAtomico(workbook, { ...opciones, loteId, lotesAnteriores })
    : { disponible: false, modo: 'indisponible', trabajadores: [], registros: [] }
  const usarAtomico = atomico.disponible && (
    clasico.importados.length === 0 || atomico.modo === 'detalle'
  )

  if (!usarAtomico) {
    return {
      ...clasico,
      loteId,
      modo: clasico.importados.length > 0 ? 'compatibilidad' : 'indisponible',
      registros: atomico.registros,
    }
  }

  const importados: ImportedWorker[] = atomico.trabajadores.map((worker) => ({
    ...worker,
    loteRemuneracion: loteId,
  }))

  const porDniActual = new Map(
    actuales.map((trabajador) => [String(trabajador.dni).trim().padStart(8, '0'), trabajador])
  )
  const trabajadores = importados.map((worker) => {
    const existente = porDniActual.get(String(worker.dni).trim().padStart(8, '0'))
    if (!existente) return worker
    return {
      ...existente,
      ...worker,
      diasTrabajados: existente.diasTrabajados,
      mesesDias: existente.mesesDias,
      mesesPresentes: Array.from(
        new Set([
          ...(worker.mesesPresentes ?? []),
          ...(existente.mesesDias ?? []),
        ])
      ),
    }
  })

  return {
    importados,
    trabajadores,
    loteId,
    modo: 'atomico',
    registros: atomico.registros,
  }
}

// ---------------------------------------------------------------------------
// EXTRAER DIAS
// ---------------------------------------------------------------------------

function nombreValido(texto: string): boolean {
  return texto.length >= 3 && /[A-Za-z]/.test(texto)
}

function esCitikold(tabla: Tabla): boolean {
  return tabla.campos.has('diasEfectivos') && tabla.campos.has('diasPosibles')
}

function citikoldPort(tabla: Tabla): ImportedWorker[] {
  const cDni = tabla.campos.get('dni')!
  const cNombres = tabla.campos.get('nombres')
  const cFull = tabla.campos.get('nombreCompleto')
  const cCodigo = tabla.campos.get('codigo')
  const cDiasPosibles = tabla.campos.get('diasPosibles')!
  const cDiasDeduccion = tabla.campos.get('diasDeduccion')
  const cDiasEfectivos = tabla.campos.get('diasEfectivos')!
  const cFechaIngreso = tabla.campos.get('fechaIngreso')
  const cFechaCese = tabla.campos.get('fechaCese')
  const cJornada = tabla.campos.get('jornada')

  let anio = 2025

  const feriados: string[] = []
  if (tabla.feriadoCol) {
    const hasta = Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 16)
    for (let dr = tabla.filaDatos; dr < hasta; dr++) {
      const fecha = celdaValorFecha(tabla.ws.getRow(dr).getCell(tabla.feriadoCol).value)
      if (fecha) {
        feriados.push(formatearFecha(fecha))
        if (feriados.length === 1) anio = fecha.getUTCFullYear()
      }
    }
  }
  const feriadosSet = new Set(feriados)

  const laborablesPorMesAnio = (fechaInicio: Date | null, fechaCese: Date | null, jornada6: boolean): number[] => {
    const bruto = new Array(12).fill(0)
    const inicio =
      fechaInicio && fechaInicio.getUTCFullYear() === anio
        ? new Date(Date.UTC(anio, fechaInicio.getUTCMonth(), fechaInicio.getUTCDate()))
        : new Date(Date.UTC(anio, 0, 1))
    const fin =
      fechaCese && fechaCese.getUTCFullYear() === anio
        ? new Date(Date.UTC(anio, fechaCese.getUTCMonth(), fechaCese.getUTCDate()))
        : new Date(Date.UTC(anio, 11, 31))
    const cur = new Date(inicio)
    while (cur <= fin) {
      const dow = cur.getUTCDay()
      const laborable = jornada6 ? dow !== 0 : dow !== 0 && dow !== 6
      if (laborable && !feriadosSet.has(cur.toISOString().slice(0, 10))) {
        bruto[cur.getUTCMonth()]++
      }
      cur.setUTCDate(cur.getUTCDate() + 1)
    }
    return bruto
  }

  const distribuirPorMes = (brutos: number[], total: number): MonthlyValues => {
    const out: MonthlyValues = { ...MESES_EN_CERO, total }
    const suma = brutos.reduce((a, b) => a + b, 0)
    if (suma <= 0) return out
    const racionales = brutos.map((b) => (total * b) / suma)
    const enteros = racionales.map(Math.floor)
    let resto = total - enteros.reduce((a, b) => a + b, 0)
    const fracciones = racionales
      .map((v, i) => [i, v - Math.floor(v)] as [number, number])
      .sort((a, b) => b[1] - a[1])
    for (let k = 0; k < resto; k++) enteros[fracciones[k]?.[0] ?? 0]++
    for (let i = 0; i < 12; i++) out[months[i]] = enteros[i]
    out.total = total
    return out
  }

  const porDni = new Map<string, ImportedWorker>()

  for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 3000); dr++) {
    const row = tabla.ws.getRow(dr)
    const dni = normalizeDni(row.getCell(cDni).value)
    if (!dni) continue

    let infoNombres: InfoNombres | null = null
    const nombreCompleto = cFull
      ? limpiarNombre(row.getCell(cFull).value)
      : cNombres
        ? limpiarNombre(row.getCell(cNombres).value)
        : ''
    if (nombreCompleto && nombreValido(nombreCompleto)) infoNombres = splitNombreCompleto(nombreCompleto)

    const diasPosibles = toNumber(row.getCell(cDiasPosibles).value) || (cCodigo ? toNumber(row.getCell(cCodigo).value) : 0)
    const diasNoLaboradosReferencia = cDiasDeduccion ? toNumber(row.getCell(cDiasDeduccion).value) : 0
    const diasEfectivosReferencia = toNumber(row.getCell(cDiasEfectivos).value) || (cCodigo ? toNumber(row.getCell(cCodigo).value) : 0)

    const fechaInicio = cFechaIngreso ? celdaValorFecha(row.getCell(cFechaIngreso).value) : null
    const fechaCese = cFechaCese ? celdaValorFecha(row.getCell(cFechaCese).value) : null

    let jornada = 5
    if (cJornada) {
      const textoJornada = String(row.getCell(cJornada).value ?? '').trim().toUpperCase()
      if (/^\d+$/.test(textoJornada)) jornada = Number(textoJornada)
      else if (textoJornada) jornada = textoJornada.includes('S') ? 6 : 5
    }

    const brutos = laborablesPorMesAnio(fechaInicio, fechaCese, jornada === 6)
    const diasTrabajados = distribuirPorMes(brutos, diasPosibles)

    const laborablesPorMes: MonthlyValues = { ...MESES_EN_CERO }
    months.forEach((m, i) => {
      laborablesPorMes[m] = brutos[i]
    })
    laborablesPorMes.total = brutos.reduce((a, b) => a + b, 0)

    const mesesDias = months.filter((mes) => diasTrabajados[mes] > 0)

    const trabajador: ImportedWorker = {
      dni,
      apellidoPaterno: infoNombres?.apellidoPaterno ?? '',
      apellidoMaterno: infoNombres?.apellidoMaterno ?? '',
      nombres: infoNombres?.nombres ?? '',
      fechaInicio: fechaInicio ? formatearFecha(fechaInicio) : '',
      fechaCese: fechaCese ? formatearFecha(fechaCese) : '',
      remuneraciones: { ...MESES_EN_CERO },
      diasTrabajados,
      incidencias: [],
      mesesDias,
      mesesPresentes: mesesDias,
      jornada,
      diasPosibles,
      diasNoLaboradosReferencia,
      diasEfectivosReferencia,
      feriados,
      laborablesPorMes,
    }

    porDni.set(dni, trabajador)
  }

  return [...porDni.values()]
}

function recogerConsolidado(tabla: Tabla): ImportedWorker[] {
  if (esCitikold(tabla)) return citikoldPort(tabla)
  return recogerConsolidadoGenerico(tabla)
}

function recogerConsolidadoGenerico(tabla: Tabla): ImportedWorker[] {
  const cDni = tabla.campos.get('dni')!
  const cDiasNoLab = tabla.campos.get('diasPosibles')
  const cDiasEfectivos = tabla.campos.get('diasEfectivos')
  const cDiasDeduccion = tabla.campos.get('diasDeduccion')
  const cDiasNeto = tabla.campos.get('diasNeto')
  const cJornada = tabla.campos.get('jornada')
  const cFechaI = tabla.campos.get('fechaI')
  const cFechaF = tabla.campos.get('fechaF')
  const cDescripcion = tabla.campos.get('descripcion')
  const cFechaIngreso = tabla.campos.get('fechaIngreso')
  const cFechaCese = tabla.campos.get('fechaCese')

  const diasEfectivosReferencia = cDiasEfectivos ? toNumber(tabla.ws.getRow(tabla.filaDatos).getCell(cDiasEfectivos).value) : 0
  const diasNoLaboradosReferencia = cDiasNoLab ? toNumber(tabla.ws.getRow(tabla.filaDatos).getCell(cDiasNoLab).value) : 0
  const semanasSize = tabla.semanas.size

  const trabajadores: ImportedWorker[] = []

  const feriados: string[] = []
  const aplicarFeriado = (v: ExcelJS.CellValue) => {
    const texto = limpiarNombre(v)
    if (nombreValido(texto)) feriados.push(texto)
  }

  if (tabla.feriadoCol) {
    const desde = Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 21)
    for (let dr = tabla.filaDatos; dr < desde; dr++) {
      aplicarFeriado(tabla.ws.getRow(dr).getCell(tabla.feriadoCol).value)
    }
  }

  for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 3000); dr++) {
    const row = tabla.ws.getRow(dr)
    const dni = normalizeDni(row.getCell(cDni).value)
    if (!dni) continue

    const nombres = nombresDesdeTabla(tabla, row)
    let infoNombres = nombres
    if (!infoNombres) {
      const nombreCompleto = caminoNombreCompleto(row, tabla)
      infoNombres = nombreCompleto ? splitNombreCompleto(nombreCompleto) : null
    }

    const fechaI = celdaValorFecha(cFechaI ? row.getCell(cFechaI).value : null) ?? celdaValorFecha(cFechaIngreso ? row.getCell(cFechaIngreso).value : null)
    const fechaF = celdaValorFecha(cFechaF ? row.getCell(cFechaF).value : null) ?? celdaValorFecha(cFechaCese ? row.getCell(cFechaCese).value : null)

    const diasDeNeto = cDiasNeto ? toNumber(row.getCell(cDiasNeto).value) : 0
    let jornada = 0
    if (cJornada) {
      const textoJornada = String(row.getCell(cJornada).value ?? '').trim()
      jornada = /^\d+$/.test(textoJornada) ? Number(textoJornada) : textoJornada ? (/S/.test(textoJornada) ? 6 : 5) : 0
    }
    const trabajador: ImportedWorker = {
      dni,
      apellidoPaterno: infoNombres?.apellidoPaterno ?? '',
      apellidoMaterno: infoNombres?.apellidoMaterno ?? '',
      nombres: infoNombres?.nombres ?? '',
      fechaInicio: fechaI ?? '',
      fechaCese: fechaF ?? '',
      remuneraciones: { ...MESES_EN_CERO },
      incidencias: [],
      diasNoLaborados: diasDeNeto,
      diasPosibles: diasNoLaboradosReferencia,
      diasEfectivos: diasEfectivosReferencia,
      diasEfectivosReferencia,
      diasNoLaboradosReferencia,
      jornada,
    }

    trabajadores.push(trabajador)

    if (cDescripcion) aplicarFeriado(row.getCell(cDescripcion).value)
  }

  const anio = inferirAnio([tabla])

  const trabajadoresConsolidados: ImportedWorker[] = []
  for (const trabajador of trabajadores) {
    const dni = trabajador.dni

    const diasEfectivosNetos = cDiasEfectivos ? Number(trabajador.diasEfectivos) : 0
    const diasDeduccion = cDiasDeduccion ? rowDiasPorCampo(tabla, dni, 'diasDeduccion') : 0

    if (cDiasNoLab) {
      const yaDias = rowDiasPorCampo(tabla, dni, 'diasPosibles')
      trabajador.diasPosibles = yaDias
      trabajador.diasEfectivos = diasEfectivosNetos
      trabajador.diasNoLaboradosReferencia = diasNoLaboradosReferencia || yaDias
      trabajador.diasEfectivosReferencia = diasEfectivosReferencia
    } else if (cDiasEfectivos) {
      trabajador.diasEfectivos = diasEfectivosNetos
      trabajador.diasNoLaboradosReferencia = diasNoLaboradosReferencia || diasEfectivosNetos
      trabajador.diasPosibles = diasNoLaboradosReferencia || diasEfectivosNetos
    } else if (cDiasDeduccion) {
      trabajador.diasNoLaborados = diasDeduccion
      trabajador.diasNoLaboradosReferencia = diasNoLaboradosReferencia || diasDeduccion
    } else if (cDiasNeto) {
      trabajador.diasNoLaborados = rowDiasPorCampo(tabla, dni, 'diasNeto')
    }

    const fechaIni = trabajador.fechaInicio instanceof Date ? trabajador.fechaInicio : null
    const fechaCes = trabajador.fechaCese instanceof Date ? trabajador.fechaCese : null
    const fechaIniDate = fechaIni ? new Date(fechaIni.getTime()) : null
    const fechaCesDate = fechaCes ? new Date(fechaCes.getTime()) : null

    const mesInicial =
      fechaIniDate && !Number.isNaN(fechaIniDate.getTime())
        ? fechaIniDate.getUTCMonth()
        : 0
    const mesFinal =
      fechaCesDate && !Number.isNaN(fechaCesDate.getTime())
        ? fechaCesDate.getUTCMonth()
        : 11

    const mesesLaborables = mesesEntre(mesInicial, mesFinal)

    const diasLaborablesBrutosPorMes = semanasSize > 0
      ? diasLaborablesBrutosPorMesAnio(anio, mesesLaborables, trabajador.diasPosibles, trabajador.diasEfectivos)
      : calcularDiasBrutosPorMes(mesInicial, mesFinal, trabajador.diasEfectivos ?? 365)

    const laborablesPorMes = distribuirTotalPorMes(diasLaborablesBrutosPorMes, mesesLaborables, Boolean(fechaIni))

    const mesesDias = months.filter((mes) => laborablesPorMes[mes] > 0)
    const mesesPresentes = Array.from(
      new Set([...(trabajador.mesesRemuneraciones ?? []), ...mesesDias])
    )

    const trabajadorFinal: ImportedWorker = {
      ...trabajador,
      mesesDias,
      mesesPresentes,
      feriados: feriados.length > 0 ? feriados : undefined,
      laborablesPorMes,
      fechaInicio: fechaIni ? formatearFecha(fechaIni) : '',
      fechaCese: fechaCes ? formatearFecha(fechaCes) : '',
    }

    if (trabajador.diasNoLaborados === undefined || trabajador.diasNoLaborados === null) {
      delete trabajadorFinal?.diasNoLaborados
    }
    if (trabajador.diasPosibles === undefined || trabajador.diasPosibles === null) {
      delete trabajadorFinal?.diasPosibles
    }
    if (trabajador.diasEfectivos === undefined || trabajador.diasEfectivos === null) {
      delete trabajadorFinal?.diasEfectivos
    }
    if (trabajador.diasNoLaboradosReferencia === undefined || trabajador.diasNoLaboradosReferencia === null) {
      delete trabajadorFinal?.diasNoLaboradosReferencia
    }
    if (trabajador.jornada === undefined || trabajador.jornada === null) {
      delete trabajadorFinal?.jornada
    }

    trabajadoresConsolidados.push(trabajadorFinal)
  }

  const trabajadoresNoRepetidos = [...new Map(trabajadoresConsolidados.map((t) => [t.dni, t])).values()]

  return trabajadoresNoRepetidos
}

function caminoNombreCompleto(row: ExcelJS.Row, tabla: Tabla): string {
  const cFull = tabla.campos.get('nombreCompleto')
  const cAp = tabla.campos.get('apellidoPaterno')
  const cAm = tabla.campos.get('apellidoMaterno')
  const cNom = tabla.campos.get('nombres')
  if (cFull) return limpiarNombre(row.getCell(cFull).value)
  if (cAp) {
    return [
      limpiarNombre(row.getCell(cAp).value),
      limpiarNombre(cAm ? row.getCell(cAm).value : ''),
      limpiarNombre(cNom ? row.getCell(cNom).value : ''),
    ].filter(Boolean).join(' ')
  }
  return ''
}

function rowDiasPorCampo(tabla: Tabla, dni: string, campo: 'diasDeduccion' | 'diasPosibles' | 'diasNeto'): number {
  const col = tabla.campos.get(campo)
  if (!col) return 0
  for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 3000); dr++) {
    const row = tabla.ws.getRow(dr)
    if (normalizeDni(row.getCell(tabla.campos.get('dni')!).value) === dni) {
      return toNumber(row.getCell(col).value)
    }
  }
  return 0
}

function mesesEntre(inicial: number, final: number): string[] {
  const meses: string[] = []
  for (let m = inicial; m <= final; m++) meses.push(months[m])
  return meses
}

function calcularDiasBrutosPorMes(inicial: number, final: number, total: number): number[] {
  if (total <= 0) return []
  const meses = final - inicial + 1
  if (meses <= 0) return []
  const base = Math.floor(total / meses)
  const resto = total % meses
  const resultado = new Array(meses).fill(base)
  for (let i = 0; i < resto; i++) resultado[i] += 1
  return resultado
}

function diasLaborablesBrutosPorMesAnio(
  anio: number,
  mesesSolicitados: string[],
  diasPosibles?: number,
  diasEfectivos?: number
): number[] {
  const laborablesPorMes = laborablesDelAnio(anio)

  const resultados: number[] = []
  for (const mes of mesesSolicitados) {
    resultados.push(laborablesPorMes[mes] ?? 0)
  }

  const totalDias = resultados.reduce((a, b) => a + b, 0)
  if (totalDias === 0) return resultados

  const ajuste = (diasPosibles ?? 0) - (diasEfectivos ?? 0)
  const ajustePorMes = ajuste / resultados.length

  return resultados.map((d) => {
    const cantidad = d + ajustePorMes
    return cantidad
  })
}

function laborablesDelAnio(anio: number): Record<string, number> {
  const resultados: Record<string, number> = months.reduce((acc, m) => {
    acc[m] = 0
    return acc
  }, {} as Record<string, number>)

  const inicio = new Date(Date.UTC(anio, 0, 1))
  const fin = new Date(Date.UTC(anio, 11, 31))

  for (let d = new Date(inicio); d <= fin; d.setUTCDate(d.getUTCDate() + 1)) {
    const dia = d.getUTCDay()
    if (dia === 0 || dia === 6) continue
    resultados[months[d.getUTCMonth()]]++
  }

  return resultados
}

function distribuirTotalPorMes(dias: number[], meses: string[], conFechas: boolean): MonthlyValues {
  const valor: MonthlyValues = { ...MESES_EN_CERO }
  if (!dias || dias.length === 0) return valor

  const suma = dias.reduce((a, b) => a + b, 0)
  let restoRestante = suma

  const mesesNombres =
    Array.isArray(meses) && meses.length > 0 && months.includes(meses[0] as (typeof months)[number])
      ? meses
      : months

  const reparto = dias.slice(0, mesesNombres.length)

  for (let i = 0; i < mesesNombres.length; i++) {
    const nombre = mesesNombres[i] as (typeof months)[number]
    const bruto = reparto[i] ?? 0
    if (reparto.length === 1) {
      valor[nombre] = conFechas ? Math.round(reparto[0]) : reparto[0]
      restoRestante = 0
    } else {
      const cuota = Math.floor(bruto)
      valor[nombre] = cuota
      restoRestante -= cuota
    }
  }

  if (reparto.length > 1 && restoRestante > 0) {
    for (let i = 0; i < mesesNombres.length && restoRestante > 0; i++) {
      valor[mesesNombres[i] as (typeof months)[number]] += 1
      restoRestante--
    }
  }

  valor.total = valor.enero + valor.febrero + valor.marzo + valor.abril + valor.mayo + valor.junio + valor.julio + valor.agosto + valor.septiembre + valor.octubre + valor.noviembre + valor.diciembre
  return valor
}

function celdaValorFecha(value: ExcelJS.CellValue): Date | null {
  if (!value) return null
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value
  if (typeof value === 'number' && Number.isFinite(value) && value > 20000 && value < 80000) {
    return new Date(Date.UTC(1899, 11, 30) + value * 86_400_000)
  }
  if (typeof value === 'string') {
    const texto = value.trim()
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(texto)
    if (iso) {
      const fecha = new Date(Date.UTC(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3])))
      return Number.isNaN(fecha.getTime()) ? null : fecha
    }
    const fecha = new Date(texto)
    return Number.isNaN(fecha.getTime()) ? null : fecha
  }
  if (typeof value === 'object' && value !== null) {
    if ('result' in value) return celdaValorFecha(value.result as ExcelJS.CellValue)
    if ('text' in value && typeof value.text === 'string') return celdaValorFecha(value.text)
  }
  return null
}

function formatearFecha(fecha: Date | null): string {
  if (!fecha || Number.isNaN(fecha.getTime())) return ''
  const anio = fecha.getUTCFullYear()
  const mes = String(fecha.getUTCMonth() + 1).padStart(2, '0')
  const dia = String(fecha.getUTCDate()).padStart(2, '0')
  return `${anio}-${mes}-${dia}`
}

function proyectarIncidencias(
  porDni: Map<string, IncidenciaNormalizada[]>
): Map<string, IncidenciaImportada[]> {
  const resultado = new Map<string, IncidenciaImportada[]>()
  for (const [dni, registros] of porDni) {
    resultado.set(
      dni,
      registros.map((registro) => ({
        dni: registro.dni,
        codigo: registro.codigo,
        descripcion: registro.descripcion,
        cantidadDias: registro.dias ?? 0,
        ...(registro.diasNeto !== null ? { diasNeto: registro.diasNeto } : {}),
        fechaInicio: registro.fechaInicio,
        fechaFin: registro.fechaFin,
        periodo: registro.periodo,
        estado: registro.estado,
        motivo: registro.motivo,
        tipo: registro.tipo,
        confianza: registro.confianza,
        origen: registro.origen,
      }))
    )
  }
  return resultado
}

function mergeConsolidados(
  consolidados: Tabla[],
  incidencias: Map<string, IncidenciaImportada[]> = new Map(),
  nombresDesdeIncidencias: Map<string, InfoNombres> = new Map(),
): ImportedWorker[] {
  if (consolidados.length === 0 && incidencias.size === 0) return []

  const porDni = new Map<string, ImportedWorker>()

  for (const tabla of consolidados) {
    for (const trabajador of recogerConsolidado(tabla)) {
      const existente = porDni.get(trabajador.dni)

      if (!existente) {
        porDni.set(trabajador.dni, trabajador)
        continue
      }

      if (!existente.apellidoPaterno && trabajador.apellidoPaterno) existente.apellidoPaterno = trabajador.apellidoPaterno
      if (!existente.apellidoMaterno && trabajador.apellidoMaterno) existente.apellidoMaterno = trabajador.apellidoMaterno
      if (!existente.nombres && trabajador.nombres) existente.nombres = trabajador.nombres
      if (!existente.fechaInicio && trabajador.fechaInicio) existente.fechaInicio = trabajador.fechaInicio
      if (!existente.fechaCese && trabajador.fechaCese) existente.fechaCese = trabajador.fechaCese
      if (trabajador.diasNoLaborados !== undefined && trabajador.diasNoLaborados !== null) existente.diasNoLaborados = trabajador.diasNoLaborados
      if (trabajador.diasPosibles !== undefined && existente.diasPosibles === undefined) existente.diasPosibles = trabajador.diasPosibles
      if (trabajador.diasEfectivos !== undefined && existente.diasEfectivos === undefined) existente.diasEfectivos = trabajador.diasEfectivos
      if (trabajador.diasNoLaboradosReferencia !== undefined) existente.diasNoLaboradosReferencia = trabajador.diasNoLaboradosReferencia
      if (trabajador.diasEfectivosReferencia !== undefined) existente.diasEfectivosReferencia = trabajador.diasEfectivosReferencia
      if (trabajador.jornada) existente.jornada = trabajador.jornada
      if (trabajador.feriados && existente.feriados === undefined) existente.feriados = trabajador.feriados
      if (trabajador.laborablesPorMes && existente.laborablesPorMes === undefined) existente.laborablesPorMes = trabajador.laborablesPorMes
      if (trabajador.diasTrabajados && existente.diasTrabajados === undefined) existente.diasTrabajados = trabajador.diasTrabajados
      if ((trabajador.incidencias?.length ?? 0) > 0) {
        existente.incidencias = [...(existente.incidencias ?? []), ...(trabajador.incidencias ?? [])]
      }
    }
  }

  for (const trabajador of porDni.values()) {
    const nombres = nombresDesdeIncidencias.get(trabajador.dni)
    if (nombres) {
      trabajador.apellidoPaterno = nombres.apellidoPaterno || trabajador.apellidoPaterno
      trabajador.apellidoMaterno = nombres.apellidoMaterno || trabajador.apellidoMaterno
      trabajador.nombres = nombres.nombres || trabajador.nombres
    }
    const lista = incidencias.get(trabajador.dni)
    if (lista?.length) trabajador.incidencias = [...(trabajador.incidencias ?? []), ...lista]
  }

  for (const [dni, lista] of incidencias) {
    if (porDni.has(dni)) continue
    const nombres = nombresDesdeIncidencias.get(dni)
    porDni.set(dni, {
      dni,
      apellidoPaterno: nombres?.apellidoPaterno ?? '',
      apellidoMaterno: nombres?.apellidoMaterno ?? '',
      nombres: nombres?.nombres ?? '',
      fechaInicio: '',
      fechaCese: '',
      remuneraciones: { ...MESES_EN_CERO },
      incidencias: [...lista],
      mesesPresentes: [],
    })
  }

  return [...porDni.values()]
}

// Encabezado de una columna de una tabla detectada, listo para trazabilidad.
function normalizarTextoDeColumna(tabla: Tabla, col: number): string {
  return normalizeHeader(tabla.ws.getRow(tabla.filaEncabezado).getCell(col).value)
}

// Quita los rótulos de total y de unidad ("TOTAL DIAS NO LABORADOS" ->
// "DIAS NO LABORADOS") para que el texto conserve solo la naturaleza del dato.
function limpiarRotuloIncidencia(texto: string): string {
  return texto
    .replace(/^(TOTAL|TOT\.?|SUMA|ACUMULADO)\b\.?\s*/i, '')
    .replace(/^(N[°º.]?\s*)?DIAS?\b\s*(NO\s+)?(LABORADOS?|NO\s+LABORADOS?)?\s*(DEL?\s*)?(A[OÑ]O|EJERCICIO|PERIODO)?\s*/i, '')
    .replace(/\s+/g, ' ')
    .trim()
}

function tituloLegible(texto: string): string {
  if (!texto) return texto
  return texto.charAt(0).toUpperCase() + texto.slice(1).toLowerCase()
}

/**
 * Rótulo y tipo de las incidencias de una matriz mensual de días no laborados.
 * Se deducen de lo que rotula la propia tabla (columna de total y nombre de la
 * hoja), nunca de un caso concreto: si el archivo no rotula nada identificable
 * se conserva la etiqueta genérica de falta.
 */
function rotuloTablaMensual(tabla: Tabla): { descripcion: string; tipo: TipoIncidencia } {
  const candidatos = [
    tabla.totalCol ? limpiarNombre(tabla.ws.getRow(tabla.filaEncabezado).getCell(tabla.totalCol).value) : '',
    tabla.ws.name,
  ]
    .map(limpiarRotuloIncidencia)
    .filter(Boolean)

  for (const tipo of ['cese', 'suspension', 'licencia_medica', 'vacaciones', 'permiso'] as TipoIncidencia[]) {
    const encontrado = candidatos.find((texto) => clasificarTipo(texto) === tipo)
    if (encontrado) return { descripcion: tituloLegible(encontrado), tipo }
  }

  const falta = candidatos.find((texto) => clasificarTipo(texto) === 'falta')
  if (falta) return { descripcion: tituloLegible(falta), tipo: 'falta' }

  return { descripcion: 'Faltas', tipo: 'falta' }
}

function procesoMensual(
  mensuales: Tabla[],
  incidencias: Map<string, IncidenciaImportada[]> = new Map(),
  nombresDesdeIncidencias: Map<string, InfoNombres> = new Map(),
  anioEjercicio: number | null = null,
): ImportedWorker[] {
  if (mensuales.length === 0) {
    // Sin cuadro mensual de días, un trabajador que solo aparece en la tabla de
    // incidencias igual se registra: queda visible en Validación con sus
    // incidencias y sin días, en vez de desaparecer.
    return [...incidencias.entries()].map(([dni, lista]) => {
      const nombres = nombresDesdeIncidencias.get(dni)
      return {
        dni,
        apellidoPaterno: nombres?.apellidoPaterno ?? '',
        apellidoMaterno: nombres?.apellidoMaterno ?? '',
        nombres: nombres?.nombres ?? '',
        fechaInicio: '',
        fechaCese: '',
        remuneraciones: { ...MESES_EN_CERO },
        incidencias: lista,
        mesesPresentes: [],
      }
    })
  }


  const estadisticas = mensuales.map((tabla) => {
    let filas = 0
    let suma = 0
    const cDni = tabla.campos.get('dni')!
    for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 5000); dr++) {
      const row = tabla.ws.getRow(dr)
      if (!normalizeDni(row.getCell(cDni).value)) continue
      filas++
      suma += calcularTotalSuma(tabla, row)
    }
    return { tabla, filas, suma }
  })

  estadisticas.sort((a, b) => b.filas - a.filas || b.suma - a.suma)
  const principal = estadisticas[0].tabla

  const nombresPorDni = new Map<string, InfoNombres>()
  const incidenciasPorDni = new Map<string, IncidenciaImportada[]>(incidencias)
  const trabajadoresDias = new Map<string, ImportedWorker>()

  const cDniPrincipal = principal.campos.get('dni')!

  for (let dr = principal.filaDatos; dr < Math.min(principal.ws.rowCount + 1, principal.filaDatos + 5000); dr++) {
    const row = principal.ws.getRow(dr)
    const dni = normalizeDni(row.getCell(cDniPrincipal).value)
    if (!dni) continue

    const nombres = nombresDesdeTabla(principal, row)
    if (nombres && (nombres.apellidoPaterno || nombres.apellidoMaterno || nombres.nombres)) {
      nombresPorDni.set(dni, nombres)
    }

    const dias = { ...MESES_EN_CERO }
    for (const mes of months) {
      dias[mes] = toNumber(row.getCell(principal.semanas.get(mes)!).value)
    }
    dias.total = (principal.totalCol ? toNumber(row.getCell(principal.totalCol).value) : 0) || calcularTotalSuma(principal, row)

    const trabajador: ImportedWorker = {
      dni,
      apellidoPaterno: '',
      apellidoMaterno: '',
      nombres: '',
      fechaInicio: principal.campos.get('fechaIngreso')
        ? formatearFecha(celdaValorFecha(row.getCell(principal.campos.get('fechaIngreso')!).value))
        : '',
      fechaCese: principal.campos.get('fechaCese')
        ? formatearFecha(celdaValorFecha(row.getCell(principal.campos.get('fechaCese')!).value))
        : '',
      remuneraciones: { ...MESES_EN_CERO },
      diasTrabajados: dias,
      ...(mensuales.length > 1 ? { diasPosibles: dias.total } : {}),
      mesesDias: months.filter((mes) => dias[mes] > 0),
      incidencias: [],
    }

    trabajadoresDias.set(dni, trabajador)
  }

  const tablasNombres = [...mensuales]
  for (const tabla of tablasNombres) {
    const cDni = tabla.campos.get('dni')!
    for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 5000); dr++) {
      const row = tabla.ws.getRow(dr)
      const dni = normalizeDni(row.getCell(cDni).value)
      if (!dni) continue

      const nombres = nombresDesdeTabla(tabla, row)
      if (!nombres) continue

      const actual = nombresPorDni.get(dni)
      if (!actual) {
        if (nombres.apellidoPaterno || nombres.apellidoMaterno || nombres.nombres) {
          nombresPorDni.set(dni, nombres)
        }
        continue
      }

      if (!actual.apellidoPaterno && nombres.apellidoPaterno) actual.apellidoPaterno = nombres.apellidoPaterno
      if (!actual.apellidoMaterno && nombres.apellidoMaterno) actual.apellidoMaterno = nombres.apellidoMaterno
      if (!actual.nombres && nombres.nombres) actual.nombres = nombres.nombres
    }
  }

  for (const [dni, nombres] of nombresDesdeIncidencias) {
    const actual = nombresPorDni.get(dni)
    if (!actual) {
      nombresPorDni.set(dni, nombres)
      continue
    }
    if (!actual.apellidoPaterno && nombres.apellidoPaterno) actual.apellidoPaterno = nombres.apellidoPaterno
    if (!actual.apellidoMaterno && nombres.apellidoMaterno) actual.apellidoMaterno = nombres.apellidoMaterno
    if (!actual.nombres && nombres.nombres) actual.nombres = nombres.nombres
  }

  for (let idx = 1; idx < estadisticas.length; idx++) {
    const tabla = estadisticas[idx].tabla
    const cDni = tabla.campos.get('dni')!
    // La tabla secundaria es una matriz mensual de días no laborados: cada mes
    // con valor es una incidencia y conserva su período. El tipo se deduce de
    // lo que rotula la propia tabla; solo si no hay rotulo identificable se usa
    // la etiqueta genérica.
    const rotulo = rotuloTablaMensual(tabla)
    for (let dr = tabla.filaDatos; dr < Math.min(tabla.ws.rowCount + 1, tabla.filaDatos + 5000); dr++) {
      const row = tabla.ws.getRow(dr)
      const dni = normalizeDni(row.getCell(cDni).value)
      if (!dni) continue
      for (const mes of months) {
        const col = tabla.semanas.get(mes)
        if (col === undefined) continue
        const valor = toNumber(row.getCell(col).value)
        if (valor <= 0) continue
        const numeroMes = months.indexOf(mes) + 1
        const lista = incidenciasPorDni.get(dni) ?? []
        lista.push({
          dni,
          codigo: '',
          descripcion: rotulo.descripcion,
          cantidadDias: valor,
          fechaInicio: null,
          fechaFin: null,
          periodo: {
            texto: etiquetaMes(numeroMes),
            anio: anioEjercicio,
            mes: numeroMes,
          },
          estado: 'identificada',
          tipo: rotulo.tipo,
          confianza: 'alta',
          origen: [
            {
              hoja: tabla.ws.name,
              fila: dr,
              columna: col,
              encabezado: normalizarTextoDeColumna(tabla, col),
              valor: String(valor),
            },
          ],
        })
        incidenciasPorDni.set(dni, lista)
      }
    }
  }

  return [...trabajadoresDias.values()].map((trabajador) => {
    const nombres = nombresPorDni.get(trabajador.dni)
    const delTrabajador = incidenciasPorDni.get(trabajador.dni) ?? []

    return {
      ...trabajador,
      apellidoPaterno: nombres?.apellidoPaterno ?? trabajador.apellidoPaterno,
      apellidoMaterno: nombres?.apellidoMaterno ?? trabajador.apellidoMaterno,
      nombres: nombres?.nombres ?? trabajador.nombres,
      incidencias: delTrabajador,
      mesesPresentes: [...(trabajador.mesesDias ?? [])],
    }
  })
}

export type ExtraccionDias = { trabajadores: ImportedWorker[]; diasImportados: number }

export function extraerDias(workbook: ExcelJS.Workbook, actuales: ImportedWorker[]): ExtraccionDias {
  const tablas = detectarTablas(workbook)
  const consolidados = tablas.filter((t) => t.tipo === 'consolidado')
  const anioEjercicio = inferirAnioEjercicio(workbook)

  // Las incidencias se identifican por semántica (encabezado + contenido) en un
  // motor aparte: no se buscan columnas "libres" ni se heredan datos de otras fuentes.
  const detectingIncidencias = extraerIncidencias(workbook, { loteId: 'dias' })
  const incidencias = proyectarIncidencias(detectingIncidencias.porDni)

  let trabajadores: ImportedWorker[]

  if (consolidados.length > 0) {
    trabajadores = mergeConsolidados(consolidados, incidencias, detectingIncidencias.nombres)
  } else {
    const mensuales = tablas
      .filter((t) => t.tipo === 'mensual')
      .filter((t) => clasificarMensual(t) === 'diasMensuales')
    trabajadores = procesoMensual(mensuales, incidencias, detectingIncidencias.nombres, anioEjercicio)
  }

  const diasImportados = trabajadores.length

  const porDniActuales = new Map(
    actuales.map((t) => [String(t.dni).trim().padStart(8, '0'), t])
  )

  const unificados = actuales.map((existente) => {
    const clave = String(existente.dni).trim().padStart(8, '0')
    const trabajador = trabajadores.find((t) => t.dni === clave)
    if (!trabajador) return existente
    const diasTrabajados = trabajador.diasTrabajados ?? existente.diasTrabajados
    return {
      ...existente,
      ...trabajador,
      remuneraciones: existente.remuneraciones ?? trabajador.remuneraciones,
      remuneracionIncompleta: existente.remuneracionIncompleta,
      mesesRemuneracionDisponibles: existente.mesesRemuneracionDisponibles,
      remuneracionResumen: existente.remuneracionResumen,
      loteRemuneracion: existente.loteRemuneracion,
      apellidoPaterno: existente.apellidoPaterno || trabajador.apellidoPaterno,
      apellidoMaterno: existente.apellidoMaterno || trabajador.apellidoMaterno,
      nombres: existente.nombres || trabajador.nombres,
      fechaInicio: trabajador.fechaInicio || existente.fechaInicio,
      fechaCese: trabajador.fechaCese || existente.fechaCese,
      // La permanencia se consolida como unión de lo que aporta cada fuente, y
      // los meses trabajados se derivan siempre de los días registrados.
      mesesDias: months.filter(
        (mes) =>
          (trabajador.mesesDias ?? []).includes(mes) || (existente.mesesDias ?? []).includes(mes)
      ),
      mesesPresentes: mesesTrabajados({
        diasTrabajados,
        mesesDias: months.filter(
          (mes) =>
            (trabajador.mesesDias ?? []).includes(mes) || (existente.mesesDias ?? []).includes(mes)
        ),
        mesesPresentes: [
          ...(trabajador.mesesPresentes ?? []),
          ...(existente.mesesPresentes ?? []),
        ],
      }),
    }
  })

  const nuevos = trabajadores.filter((t) => !porDniActuales.has(t.dni))

  return { trabajadores: [...unificados, ...nuevos], diasImportados }
}