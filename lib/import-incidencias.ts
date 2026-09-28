import * as ExcelJS from 'exceljs'

// ---------------------------------------------------------------------------
// MOTOR GENERAL DE INCIDENCIAS
//
// Identifica regiones de incidencia por semántica (encabezado + contenido),
// nunca por posiciones fijas ni por valores concretos de un archivo.
//
// Reglas del modelo:
// 1. Una columna solo aporta días o fechas si su encabezado o su contenido la
//    identifican como tal. Un importe nunca se toma como fecha ni como días.
// 2. Las columnas de ingreso/cese son datos del trabajador, no del rango de la
//    incidencia: no habilitan una región ni alimentan fechaInicio/fechaFin.
// 3. Las columnas de mes o de período tampoco son fechas de incidencia.
// 4. Si un registro no se puede identificar con seguridad, se emite con
//    estado "no_identificada" y un motivo, para revisión en Validación; nunca
//    se completa con datos de otra fuente.
// ---------------------------------------------------------------------------

export type EstadoIncidencia = 'identificada' | 'no_identificada'
export type Confianza = 'alta' | 'media'

export type TipoIncidencia =
  | 'falta'
  | 'vacaciones'
  | 'licencia_medica'
  | 'permiso'
  | 'suspension'
  | 'cese'
  | 'otro'

export type OrigenCampo = {
  hoja: string
  fila: number
  columna: number
  encabezado: string
  valor: string
}

// Período al que pertenece una incidencia. El archivo puede deliverar un año
// y un mes ("2025-02", "FEBRERO 2025"), solo un mes (columna de meses) o solo
// un texto libre. Nunca se completa con datos de otra fuente.
export type PeriodoIncidencia = {
  texto: string
  anio: number | null
  mes: number | null
}

export type IncidenciaNormalizada = {
  id: string
  loteId: string
  dni: string
  tipo: TipoIncidencia
  descripcion: string
  codigo: string
  dias: number | null
  diasNeto: number | null
  fechaInicio: Date | null
  fechaFin: Date | null
  periodo: PeriodoIncidencia | null
  estado: EstadoIncidencia
  motivo: string
  confianza: Confianza
  origen: OrigenCampo[]
}

export type ResumenRegion = {
  hoja: string
  filaEncabezado: number
  columnas: string[]
  registros: number
  identificados: number
  pendientes: number
  motivo: string
}

export type ResultadoIncidencias = {
  registros: IncidenciaNormalizada[]
  porDni: Map<string, IncidenciaNormalizada[]>
  nombres: Map<string, { apellidoPaterno: string; apellidoMaterno: string; nombres: string }>
  regiones: ResumenRegion[]
}

type Opciones = {
  loteId?: string
  maxFilasPorRegion?: number
  maxPendientesPorRegion?: number
}

// ---------------------------------------------------------------------------
// Normalización de texto, documentos y fechas
// ---------------------------------------------------------------------------

const MESES_TEXTO = [
  'ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
  'JULIO', 'AGOSTO', 'SETIEMBRE', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE',
  'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SET', 'SEP', 'OCT', 'NOV', 'DIC',
]

export function normalizarTexto(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return ''
  if (typeof value === 'object') {
    if ('text' in value) return normalizarTexto(value.text as ExcelJS.CellValue)
    if ('result' in value) return normalizarTexto(value.result as ExcelJS.CellValue)
    if ('richText' in value) {
      const rich = (value as { richText?: { text?: string }[] }).richText
      return normalizarTexto((rich ?? []).map((r) => r.text ?? '').join(' ') as ExcelJS.CellValue)
    }
    return ''
  }
  return String(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase()
}

const compacto = (texto: string): string => texto.replace(/[^A-Z0-9]/g, '')

export function normalizarDocumento(value: ExcelJS.CellValue): string | null {
  const texto = String(valorATextoPlano(value) ?? '').trim().replace(/\.0$/, '')
  if (!/^\d+$/.test(texto)) return null
  if (texto.length === 9) return texto
  if (texto.length === 7 || texto.length === 8) return texto.padStart(8, '0')
  return null
}

function valorATextoPlano(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''
  if (value instanceof Date) return ''
  if (typeof value === 'object') {
    if ('result' in value) return valorATextoPlano(value.result as ExcelJS.CellValue)
    if ('text' in value) return valorATextoPlano(value.text as ExcelJS.CellValue)
    return ''
  }
  return String(value)
}

function numeroDe(value: ExcelJS.CellValue): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (value instanceof Date) return null
  const texto = valorATextoPlano(value).trim()
  if (!texto) return null
  if (/^-?\d{1,3}(\.\d{3})+$/.test(texto)) return Number(texto.replace(/\./g, ''))
  if (/^-?\d+,\d+$/.test(texto)) return Number(texto.replace(',', '.'))
  if (!/^-?\d+(\.\d+)?$/.test(texto)) return null
  const n = Number(texto)
  return Number.isFinite(n) ? n : null
}

function valorVacio(value: ExcelJS.CellValue): boolean {
  if (value === null || value === undefined) return true
  if (typeof value === 'object' && !(value instanceof Date)) {
    const tieneTexto = 'text' in value && String((value as { text?: string }).text ?? '').trim() !== ''
    const tieneResultado = 'result' in value && (value as { result?: unknown }).result !== undefined && (value as { result?: unknown }).result !== null
    if (!tieneTexto && !tieneResultado) return true
  }
  if (typeof value === 'string') return value.trim() === ''
  return false
}

const EPOCH_EXCEL = Date.UTC(1899, 11, 30)

export function leerFecha(value: ExcelJS.CellValue): Date | null {
  if (value === null || value === undefined) return null
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value

  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value < 20000 || value > 80000) return null
    return new Date(EPOCH_EXCEL + value * 86_400_000)
  }

  const texto = valorATextoPlano(value).trim()
  if (!texto) return null

  // Un año-mes es un período, no un día: no se toma como fecha de incidencia.
  const iso = /^(\d{4})[-/](\d{1,2})(?:[-/](\d{1,2}))?$/.exec(texto)
  if (iso) {
    if (!iso[3]) return null
    return fechaValida(Number(iso[1]), Number(iso[2]), Number(iso[3]))
  }

  const numerica = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})/.exec(texto)
  if (numerica) {
    const primero = Number(numerica[1])
    const segundo = Number(numerica[2])
    let anio = Number(numerica[3])
    if (anio < 100) anio += anio >= 70 ? 1900 : 2000
    const dia = primero > 12 ? primero : segundo > 12 ? segundo : primero
    const mes = primero > 12 ? segundo : segundo > 12 ? primero : primero
    return fechaValida(anio, mes, dia)
  }

  const conMesEnLetras = new RegExp(
    `^(\\d{1,2})\\s*(DE\\s*)?(${MESES_TEXTO.join('|')})[A-Z]*\\.?\\s*(DE\\s*)?(\\d{2,4})$`,
  ).exec(texto)
  if (conMesEnLetras) {
    const dia = Number(conMesEnLetras[1])
    const mes = mesDesdeTexto(conMesEnLetras[3])
    let anio = Number(conMesEnLetras[4])
    if (anio < 100) anio += anio >= 70 ? 1900 : 2000
    if (mes === null) return null
    return fechaValida(anio, mes, dia)
  }

  return null
}

function fechaValida(anio: number, mes: number, dia: number): Date | null {
  if (!Number.isInteger(anio) || !Number.isInteger(mes) || !Number.isInteger(dia)) return null
  if (mes < 1 || mes > 12 || dia < 1 || dia > 31) return null
  const fecha = new Date(Date.UTC(anio, mes - 1, dia))
  if (fecha.getUTCFullYear() !== anio || fecha.getUTCMonth() !== mes - 1 || fecha.getUTCDate() !== dia) return null
  return fecha
}

const fechaIso = (fecha: Date | null): string => (fecha ? fecha.toISOString().slice(0, 10) : '')

const ETIQUETAS_MES = [
  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
  'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
]

// Nombre de mes -> número. SETIEMBRE y SEPTIEMBRE son el mismo mes (9).
const MESES_POR_NOMBRE: Record<string, number> = {
  ENERO: 1, FEBRERO: 2, MARZO: 3, ABRIL: 4, MAYO: 5, JUNIO: 6,
  JULIO: 7, AGOSTO: 8, SETIEMBRE: 9, SEPTIEMBRE: 9, OCTUBRE: 10, NOVIEMBRE: 11, DICIEMBRE: 12,
  ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, SET: 9, SEP: 9, OCT: 10, NOV: 11, DIC: 12,
}

function mesDesdeTexto(texto: string): number | null {
  return MESES_POR_NOMBRE[texto.toUpperCase()] ?? null
}

function anioDesdeTexto(texto: string, admiteDosDigitos = false): number | null {
  const cuatro = /(\d{4})/.exec(texto)
  if (cuatro) {
    const anio = Number(cuatro[1])
    if (anio >= 1900 && anio <= 2200) return anio
  }
  // Un año de dos dígitos solo se acepta junto a un nombre de mes ("FEB-25"):
  // en cualquier otro texto un número suelto no es un año.
  if (!admiteDosDigitos) return null
  const dos = /(?:^|\D)(\d{2})(?:\D|$)/.exec(texto)
  if (!dos) return null
  const n = Number(dos[1])
  return n >= 70 ? 1900 + n : 2000 + n
}

// ---------------------------------------------------------------------------
// Período de la incidencia
// ---------------------------------------------------------------------------

// Lee el período tal como viene en el archivo. No inventa fechas: si el archivo
// solo aporta un mes, el período conserva el mes y el año queda pendiente.
export function leerPeriodo(value: ExcelJS.CellValue): PeriodoIncidencia | null {
  if (valorVacio(value)) return null

  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null
    return { texto: fechaIso(value), anio: value.getUTCFullYear(), mes: value.getUTCMonth() + 1 }
  }

  const texto = valorATextoPlano(value).trim()
  if (!texto) return null

  // "2025-02", "2025/02", "2025.02"
  const anioMes = /^(\d{4})\s*[-/.]\s*(\d{1,2})$/.exec(texto)
  if (anioMes) {
    const mes = Number(anioMes[2])
    if (mes >= 1 && mes <= 12) return { texto, anio: Number(anioMes[1]), mes }
  }

  // "02/2025", "02-2025"
  const mesAnio = /^(\d{1,2})\s*[-/.]\s*(\d{4})$/.exec(texto)
  if (mesAnio) {
    const mes = Number(mesAnio[1])
    if (mes >= 1 && mes <= 12) return { texto, anio: Number(mesAnio[2]), mes }
  }

  // "FEBRERO", "FEB-25", "FEBRERO 2025", "2025 FEBRERO"
  const conMes = new RegExp(`(${MESES_TEXTO.join('|')})`).exec(texto.toUpperCase())
  if (conMes) {
    return { texto, anio: anioDesdeTexto(texto, true), mes: mesDesdeTexto(conMes[1]) }
  }

  // Una fecha completa en una columna de período pertenece a ese mes.
  const fecha = leerFecha(texto)
  if (fecha) {
    return { texto, anio: fecha.getUTCFullYear(), mes: fecha.getUTCMonth() + 1 }
  }

  // Cualquier otro texto se conserva literal, sin interpretarlo.
  return { texto, anio: anioDesdeTexto(texto), mes: null }
}

// Etiqueta lista para Validación. El año solo se muestra si el archivo lo
// aporta o si quien la llama conoce el año del ejercicio.
export function formatearPeriodo(
  periodo: PeriodoIncidencia | null | undefined,
  anioPorDefecto: number | null = null,
): string {
  if (!periodo) return ''
  const anio = periodo.anio ?? anioPorDefecto
  if (periodo.mes) return anio ? `${ETIQUETAS_MES[periodo.mes - 1]} ${anio}` : ETIQUETAS_MES[periodo.mes - 1]
  if (periodo.anio) return String(periodo.anio)
  return periodo.texto
}

export function etiquetaMes(mes: number | null | undefined): string {
  return mes && mes >= 1 && mes <= 12 ? ETIQUETAS_MES[mes - 1] : ''
}

function esTextoMes(texto: string): boolean {
  return MESES_TEXTO.includes(texto.toUpperCase())
}

// El texto de la incidencia se conserva tal como viene en el archivo (solo se
// limpian artefactos de formato); para comparar y clasificar se usa la forma
// normalizada sin tildes.
function limpiarTextoValor(value: ExcelJS.CellValue): string {
  const texto = valorATextoPlano(value).trim()
  return texto
    .replace(/\s+(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+[A-Za-z]{3}\s+\d{1,2}\s+\d{4}.*$/gi, '')
    .trim()
}

// ---------------------------------------------------------------------------
// Semántica de encabezados
// ---------------------------------------------------------------------------

const RE_DOCUMENTO = /\b(DNI|CE\s?D|CEDULA|CED|DOCUMENTO)\b/ 
const RE_DOCUMENTO_NUMERO = /(^|\s)(NRO|NUM|NUMERO|N)\s*\.?\s*(DE\s*)?(DOC|DOCUMENTO|IDENT|CED)/
const RE_APELLIDO_PATERNO = /APELLIDO\s*(PATERNO|1ERO|1RO|PRIMERO)\b|\bPATERNO\b/
const RE_APELLIDO_MATERNO = /APELLIDO\s*(MATERNO|2DO|2DO|SEGUNDO)\b|\bMATERNO\b/
const RE_NOMBRES = /^NOMBRE(S)?$|NOMBRE(S)?\s+Y\s+APELLIDO|APELLIDOS?\s+Y\s+NOMBRE/
const RE_NOMBRE_COMPLETO = /^APELLIDOS?\s+Y\s+NOMBRES?$|^NOMBRE(S)?\s+COMPLETO$|^NOMBRE(S)?\s+APELLIDO(S)?$/
const RE_CODIGO = /\bCOD(IGO)?\b|^COD\.?$|CLA(SVE|E)\b|ID_?INCIDENCIA|CODIGO DE INCIDENCIA/
const RE_TIPO_INCIDENCIA =
  /CONCEPTO|MOTIVO|CAUSA|DESCRIPCION|GLOSA|INCIDENCIA|ACONTECIMIENTO|EVENTO|SITUACION|CONCEPTO DE (FALTA|INCIDENCIA)|(TIPO|CLASE)\s+(DE\s+)?(LA\s+)?(INCIDENCIA|FALTA|BAJA|SUSPENSION|AUSENCIA)/
const RE_TIPO_GENERICO = /^TIPO(\s+(DE\s+)?(QUINTA|PENSION|PRIMA|DOCUMENTO|DOC|INGRESO|REMUNERACION))?(\s+DE\s+\w+)?$/
const RE_DIAS = /(^|\s)(NUM|NRO|NUMERO|CANT|CANTIDAD|CANT\.)\.?(ERO|ERO)?\s*(DE\s*)?DIAS(\b|$)|^DIAS(\b|$)|DIAS\s+(DE\s+)?(INCIDENCIA|NO\s+LABOR|FALTA|VACACION|LICENCIA)/
const RE_DIAS_TOTAL = /(TOTAL|ACUMULAD|SUMATORIA|SUMA|Σ)/
const RE_DIAS_NETO = /(DIAS\s+)?(NETOS?|NETO\b|EFECTIVOS?|CALCULADOS?|LABORABLES|COMPUTO|COMPUTADOS)/
const RE_PERIODO = /PERIODO|CICLO|PLANILLA|COMPETENCIA/
const RE_MEDIDA = /MEDIDA|UNIDAD( DE MEDIDA)?/
const RE_MONEDA = /MONEDA|DIVISA|CAPITAL/
const RE_MONTO = /MONTO|IMPORTE|VALOR|APORTE|SUELDO|REMUN|INGRESO|PAGO|DEVENGAD|BRUTO|COSTO|PRECIO|TARIFA|COMISION|RETENCION|IMPUESTO|DESCUENTO|NETO A PAGAR|HABER|CARGOS?|MONEDAS?$/
const RE_FECHA_INICIO = /(FECHA|F\.?)\s*(DE\s*)?(INICIO|INI\.?|DESDE|ENTRADA|INGRESO|CONTRATO|INICIO DE)|(^|\s)DESDE(\s|$)|^F\.?\s*INI|^INICIO$/
const RE_FECHA_FIN = /(FECHA|F\.?)\s*(DE\s*)?(FIN|FINAL|HASTA|TERMINO|CESE|SALIDA)|^F\.?\s*FIN|^HASTA(\s|$)|^FINAL$/
const RE_FECHA_CESE = /FECHA\s*(DE\s*)?CESE|CESE$|FECHA\s+DE\s+BAJA/
const RE_FECHA_INGRESO = /FECHA\s*(DE\s*)?(INGRESO|CONTRATO|ALTA|INCORPORACION)/

function puntajeDocumento(encabezado: string): number {
  if (!encabezado) return 0
  if (/^(TIPO|TYPE)\b/.test(encabezado)) return 0
  const c = compacto(encabezado)
  if (c === 'DNI' || c === 'NRODOCUMENTO' || c === 'NUMERODEDOCUMENTO' || c === 'DOCUMENTODEIDENTIDAD') return 100
  if (c === 'D' || c === 'DOC' || c === 'CED' || c === 'CE' || c === 'NRODOCIDE' || c === 'NDOCUMENTO') return 55
  if (RE_DOCUMENTO_NUMERO.test(encabezado)) return 90
  if (RE_DOCUMENTO.test(encabezado) && /IDENT|NUM|NRO/.test(encabezado)) return 80
  if (RE_DOCUMENTO.test(encabezado)) return 40
  if (/^DNI(\s|$)/.test(encabezado)) return 100
  return 0
}

// Vocabulario de remuneración: nunca es una incidencia.
const RE_CONCEPTO_REMUNERATIVO =
  /SUELDO|REMUNER|S\.?\s*BASIC|BASICO|BONO|BONIF|GRATIF|AFP|ONP|ESSALUD|5TA|CINQUENA|IMPUESTO|RETENC|PRESTAMO|ADELANTO|CTS|NETO A PAGAR|APORTE|APORTACION|JORNADA EXTRA|HORA EXTRA|HORAS EXTRA|ESTIMULO|PREMIUM|COMISION|VIATICO|TRANSPORTE|REFRIGERIO|MOVILIDAD|ASIG\.?\s*FAM|ASIGNACION FAMILIAR|CONSTANC|ALQUILER|COCHERA|PENDIENTE DE PAGO|CAPITAL/

const RE_FALTA = /FALTA|AUSENCIA|NO ASIST|INASIST|OMISION/
const RE_VACACIONES = /VACACION|DESCANSO VACACIONAL/
const RE_LICENCIA = /LICENCIA|ENFERM|ACCIDENTE|BAJA MEDICA|DME|SALUD|PATOLOG|RESGO/
const RE_PERMISO = /PERMISO/
const RE_SUSPENSION = /SUSPENSION|SUSPEND/
const RE_CESE_INCIDENCIA = /CESE|BAJA|DESPIDO|CESION/

export function clasificarTipo(texto: string): TipoIncidencia {
  const t = normalizarTexto(texto)
  if (!t) return 'otro'
  if (RE_CESE_INCIDENCIA.test(t)) return 'cese'
  if (RE_LICENCIA.test(t)) return 'licencia_medica'
  if (RE_VACACIONES.test(t)) return 'vacaciones'
  if (RE_PERMISO.test(t)) return 'permiso'
  if (RE_SUSPENSION.test(t)) return 'suspension'
  if (RE_FALTA.test(t)) return 'falta'
  return 'otro'
}

export function esConceptoRemunerativo(texto: string): boolean {
  return RE_CONCEPTO_REMUNERATIVO.test(normalizarTexto(texto))
}

// ---------------------------------------------------------------------------
// Extensión de las hojas (con caché: algunas hojas declaran 1,048,576 filas)
// ---------------------------------------------------------------------------

type Extension = { ultimaFila: number; columnas: number }
const CACHE_EXTENSION = new WeakMap<ExcelJS.Worksheet, Extension>()

function extension(ws: ExcelJS.Worksheet): Extension {
  const cacheada = CACHE_EXTENSION.get(ws)
  if (cacheada) return cacheada
  let ultima = 0
  let columnas = 0
  ws.eachRow({ includeEmpty: false }, (row) => {
    if (row.number > ultima) ultima = row.number
    if (row.cellCount > columnas) columnas = row.cellCount
  })
  const valor: Extension = {
    ultimaFila: ultima > 0 ? ultima : (ws.rowCount || 0),
    columnas: Math.min(Math.max(columnas, 1), 2000),
  }
  CACHE_EXTENSION.set(ws, valor)
  return valor
}

// ---------------------------------------------------------------------------
// Regiones candidatas y perfiles de columna
// ---------------------------------------------------------------------------

type Region = {
  ws: ExcelJS.Worksheet
  hoja: string
  filaEncabezado: number
  filasEncabezado: number[]
  filaDatos: number
  finBloque: number
  colDocumento: number
  columnas: PerfilColumna[]
  rol: Map<number, Rol>
  aceptada: boolean
  motivo: string
}

type Rol =
  | 'documento'
  | 'nombre'
  | 'apellidoPaterno'
  | 'apellidoMaterno'
  | 'nombres'
  | 'descripcion'
  | 'codigo'
  | 'dias'
  | 'diasNeto'
  | 'fechaInicio'
  | 'fechaFin'
  | 'fechaIngreso'
  | 'fechaCese'
  | 'fechaSinClasificar'
  | 'periodo'
  | 'mes'
  | 'monto'
  | 'medida'
  | 'moneda'
  | 'otro'

type PerfilColumna = {
  col: number
  ruta: string[]
  encabezado: string
  filas: number
  fechas: number
  enterosCorto: number
  montos: number
  textos: number
  documentos: number
  tiposDocumento: number
}

const VALORES_TIPO_DOCUMENTO = new Set(['DNI', 'CE', 'C.E', 'CED', 'CEDULA', 'C.I', 'PASAPORTE', 'CARNET', 'TITULO'])

// Un encabezado puede ser un rótulo de columna o un banner de la hoja. Se
// descarta el banner (metadato, no etiqueta de columna) para no mezclarlo con
// la semántica de la columna.
function esEncabezadoDeColumna(ws: ExcelJS.Worksheet, fila: number, texto: string, columnas: number): boolean {
  if (!texto) return false
  if (texto.includes(':')) return false
  if (texto.length > 45) return false
  if (/^R\d{2}\b/.test(texto)) return false

  let repetidas = 0
  for (let col = 1; col <= columnas; col++) {
    if (normalizarTexto(ws.getRow(fila).getCell(col).value) === texto) repetidas++
  }
  if (repetidas >= 6) return false

  return true
}

function perfilColumna(ws: ExcelJS.Worksheet, col: number, ruta: string[], filaDatos: number, limite: number): PerfilColumna {
  const perfil: PerfilColumna = {
    col,
    ruta,
    encabezado: ruta.join(' '),
    filas: 0,
    fechas: 0,
    enterosCorto: 0,
    montos: 0,
    textos: 0,
    documentos: 0,
    tiposDocumento: 0,
  }
  const hasta = Math.min(limite, filaDatos + 200)
  for (let fila = filaDatos; fila <= hasta; fila++) {
    const valor = ws.getRow(fila).getCell(col).value
    if (valorVacio(valor)) continue
    perfil.filas++
    if (normalizarDocumento(valor)) {
      perfil.documentos++
      continue
    }
    if (leerFecha(valor)) {
      perfil.fechas++
      continue
    }
    const numero = numeroDe(valor)
    if (numero !== null) {
      if (Number.isInteger(numero) && numero >= 0 && numero <= 31) perfil.enterosCorto++
      if (Math.abs(numero) >= 100) perfil.montos++
      continue
    }
    const texto = normalizarTexto(valor)
    if (!texto) continue
    if (esTextoMes(texto)) continue
    if (VALORES_TIPO_DOCUMENTO.has(texto)) {
      perfil.tiposDocumento++
      continue
    }
    perfil.textos++
  }
  return perfil
}

const fraccion = (
  perfil: PerfilColumna,
  campo: 'fechas' | 'enterosCorto' | 'montos' | 'textos' | 'documentos' | 'tiposDocumento'
): number => (perfil.filas === 0 ? 0 : perfil[campo] / perfil.filas)

function rolColumna(perfil: PerfilColumna): Rol {
  const encabezado = perfil.encabezado
  const fracFechas = fraccion(perfil, 'fechas')
  const fracEnteros = fraccion(perfil, 'enterosCorto')
  const fracMontos = fraccion(perfil, 'montos')
  const fracTextos = fraccion(perfil, 'textos')
  const fracDocumentos = fraccion(perfil, 'documentos')
  const fracTiposDocumento = fraccion(perfil, 'tiposDocumento')
  const esTotal = RE_DIAS_TOTAL.test(encabezado)
  const esPago = RE_MONTO.test(encabezado)

  if (fracTiposDocumento >= 0.6) return 'otro'
  if (puntajeDocumento(encabezado) > 0 && fracDocumentos >= 0.5) return 'documento'
  // Para los nombres manda la etiqueta más específica del encabezado.
  const ultimo = perfil.ruta[perfil.ruta.length - 1] ?? encabezado
  if (RE_APELLIDO_MATERNO.test(ultimo)) return 'apellidoMaterno'
  if (RE_APELLIDO_PATERNO.test(ultimo)) return 'apellidoPaterno'
  if (RE_NOMBRE_COMPLETO.test(ultimo)) return 'nombre'
  if (RE_NOMBRES.test(ultimo)) return 'nombres'
  if (RE_MEDIDA.test(encabezado)) return 'medida'
  if (RE_MONEDA.test(encabezado) && !esPago) return 'moneda'
  if (RE_PERIODO.test(encabezado)) return 'periodo'
  if (esTextoMes(encabezado)) return 'mes'
  if (RE_FECHA_CESE.test(encabezado)) return 'fechaCese'
  if (RE_FECHA_INGRESO.test(encabezado)) return 'fechaIngreso'
  if (RE_FECHA_FIN.test(encabezado)) return 'fechaFin'
  if (RE_FECHA_INICIO.test(encabezado)) return 'fechaInicio'
  if (RE_DIAS.test(encabezado) && !esPago) {
    if (RE_DIAS_NETO.test(encabezado)) return 'diasNeto'
    if (esTotal) return 'otro'
    return 'dias'
  }
  if (RE_CODIGO.test(encabezado) && fracFechas < 0.2 && (fracTextos >= 0.5 || fracEnteros >= 0.5)) return 'codigo'
  if (RE_TIPO_INCIDENCIA.test(encabezado) && !RE_TIPO_GENERICO.test(encabezado) && !esPago && fracTextos >= 0.5) {
    return 'descripcion'
  }
  if (RE_MONTO.test(encabezado) && fracFechas < 0.5) return 'monto'

  // Sin encabezado suficiente: se decide por contenido, con reglas estrictas.
  if (!encabezado) {
    if (fracFechas >= 0.9) return 'fechaSinClasificar'
    if (fracEnteros >= 0.95 && fracMontos === 0) return 'diasNeto'
    return 'otro'
  }
  if (fracFechas >= 0.9 && !esPago) return 'fechaSinClasificar'
  return 'otro'
}

function pareceFilaEncabezado(ws: ExcelJS.Worksheet, fila: number, columnas: number): boolean {
  const row = ws.getRow(fila)
  let conTexto = 0
  let conValor = 0
  for (let col = 1; col <= columnas; col++) {
    const valor = row.getCell(col).value
    if (valorVacio(valor)) continue
    conValor++
    if (typeof valor === 'number' || valor instanceof Date) continue
    const texto = normalizarTexto(valor)
    if (texto && !/^\d+([.,]\d+)?$/.test(texto)) conTexto++
  }
  if (conValor < 3) return false
  return conTexto >= 3 && conTexto / conValor >= 0.6
}

// Una hoja puede traer varios bloques apilados; cada bloque se detecta por
// contenido (filas con documento), no por la posición de la tabla.
function construirRegiones(workbook: ExcelJS.Workbook): Region[] {
  const regiones: Region[] = []

  for (const ws of workbook.worksheets) {
    const ext = extension(ws)
    const candidatasDocumento = columnasConDocumentos(ws, ext.columnas)
    if (candidatasDocumento.length === 0) continue

    for (const bloque of bloquesDeDatos(ws, ext, candidatasDocumento)) {
      const rutas = new Map<number, string[]>()
      for (let col = 1; col <= ext.columnas; col++) {
        const ruta = rutaEncabezadoDesde(ws, col, bloque.filasEncabezado, ext.columnas)
        if (ruta.length > 0) rutas.set(col, ruta)
      }
      if (rutas.size === 0) continue

      let colDocumento: number | null = null
      let mejorPuntaje = 0
      for (const col of candidatasDocumento) {
        const ruta = rutas.get(col) ?? []
        const perfil = perfilColumna(ws, col, ruta, bloque.inicio, bloque.fin)
        const fracDocumentos = fraccion(perfil, 'documentos')
        if (fracDocumentos < 0.5) continue
        // El contenido manda: una columna de documentos vale mas que el rotulo.
        const puntaje = Math.max(puntajeDocumento(ruta.join(' ')), fracDocumentos >= 0.8 ? 60 : 50)
        if (puntaje <= mejorPuntaje) continue
        mejorPuntaje = puntaje
        colDocumento = col
      }
      if (colDocumento === null) continue

      const perfiles: PerfilColumna[] = []
      const rol = new Map<number, Rol>()
      for (let col = 1; col <= ext.columnas; col++) {
        const ruta = rutas.get(col) ?? []
        const perfil = perfilColumna(ws, col, ruta, bloque.inicio, bloque.fin)
        if (perfil.filas === 0) continue
        if (ruta.length === 0 && !columnaUtilSinEncabezado(perfil)) continue
        perfiles.push(perfil)
        rol.set(col, rolColumna(perfil))
      }

      const region: Region = {
        ws,
        hoja: ws.name,
        filaEncabezado: bloque.filasEncabezado[bloque.filasEncabezado.length - 1] ?? bloque.inicio - 1,
        filasEncabezado: bloque.filasEncabezado,
        filaDatos: bloque.inicio,
        finBloque: bloque.fin,
        colDocumento,
        columnas: perfiles,
        rol,
        aceptada: false,
        motivo: '',
      }
      evaluarRegion(region)
      regiones.push(region)
    }
  }

  return regiones
}

type BloqueDatos = { filasEncabezado: number[]; inicio: number; fin: number }

function filasEncabezadoSobre(ws: ExcelJS.Worksheet, filaInicio: number, columnas: number, maximo: number = 4): number[] {
  const filas: number[] = []
  for (let fila = filaInicio - 1; fila >= 1 && filas.length < maximo; fila--) {
    if (!pareceFilaEncabezado(ws, fila, columnas)) break
    filas.unshift(fila)
  }
  return filas
}

// Se localiza por contenido cada bloque de filas con documentos: asi el
// encabezado puede ocupar varias filas (grupos y subencabezados combinados).
function bloquesDeDatos(ws: ExcelJS.Worksheet, ext: Extension, candidatasDocumento: number[]): BloqueDatos[] {
  const bloques: BloqueDatos[] = []
  let inicio: number | null = null
  let fin = 0
  let filasEncabezado: number[] = []
  let huecos = 0

  for (let fila = 1; fila <= ext.ultimaFila; fila++) {
    const tieneDocumento = candidatasDocumento.some((col) => normalizarDocumento(ws.getRow(fila).getCell(col).value))
    if (tieneDocumento) {
      if (inicio === null) {
        inicio = fila
        filasEncabezado = filasEncabezadoSobre(ws, fila, ext.columnas)
      }
      fin = fila
      huecos = 0
      continue
    }
    if (inicio === null) continue
    huecos++
    if (huecos >= 4) {
      bloques.push({ filasEncabezado, inicio, fin })
      inicio = null
      huecos = 0
    }
  }
  if (inicio !== null) bloques.push({ filasEncabezado, inicio, fin })
  return bloques
}

function columnasConDocumentos(ws: ExcelJS.Worksheet, columnas: number): number[] {
  const candidatas: number[] = []
  for (let col = 1; col <= columnas; col++) {
    let filas = 0
    let documentos = 0
    const hasta = Math.min(extension(ws).ultimaFila, 200)
    for (let fila = 1; fila <= hasta; fila++) {
      const valor = ws.getRow(fila).getCell(col).value
      if (valorVacio(valor)) continue
      filas++
      if (normalizarDocumento(valor)) documentos++
    }
    if (filas > 0 && documentos / filas >= 0.5) candidatas.push(col)
  }
  return candidatas
}

function rutaEncabezadoDesde(ws: ExcelJS.Worksheet, col: number, filasEncabezado: number[], columnas: number): string[] {
  const ruta: string[] = []
  for (const fila of filasEncabezado) {
    const texto = normalizarTexto(ws.getRow(fila).getCell(col).value)
    if (!texto || ruta[ruta.length - 1] === texto) continue
    if (!esEncabezadoDeColumna(ws, fila, texto, columnas)) continue
    ruta.push(texto)
  }
  return ruta
}

// Una columna sin encabezado solo se considera si su contenido es inequivoco.
function columnaUtilSinEncabezado(perfil: PerfilColumna): boolean {
  if (perfil.filas === 0) return false
  return fraccion(perfil, 'fechas') >= 0.9 || (fraccion(perfil, 'enterosCorto') >= 0.95 && fraccion(perfil, 'montos') === 0)
}

function columnasConRol(region: Region, rol: Rol): PerfilColumna[] {
  return region.columnas.filter((columna) => region.rol.get(columna.col) === rol)
}

function evaluarRegion(region: Region): void {
  const meses = columnasConRol(region, 'mes').length
  const periodos = columnasConRol(region, 'periodo').length
  const montos = columnasConRol(region, 'monto').length
  const descripciones = columnasConRol(region, 'descripcion')
  const dias = columnasConRol(region, 'dias')
  const diasNeto = columnasConRol(region, 'diasNeto')
  const fechasInicio = columnasConRol(region, 'fechaInicio')
  const fechasFin = columnasConRol(region, 'fechaFin')
  const fechasSinClasificar = columnasConRol(region, 'fechaSinClasificar')

  const esTablaDeDias = meses >= 3 || (meses > 0 && dias.length > 0 && descripciones.length === 0)
  if (esTablaDeDias) {
    region.motivo = 'Tabla de días por mes: no es una fuente de incidencias.'
    return
  }

  const rangoFechas = fechasInicio.length + fechasFin.length + fechasSinClasificar.length
  const fechasSemanticas = fechasInicio.length + fechasFin.length
  const tieneTipo = descripciones.length > 0
  const tieneDimension = dias.length > 0 || diasNeto.length > 0 || rangoFechas > 0

  // Sin tipo de incidencia ni rango de fechas, una tabla de días por trabajador
  // no es una fuente de incidencias: se trata como información de días.
  if (!tieneTipo && rangoFechas === 0) {
    region.motivo = 'Tabla de días por trabajador (sin tipo ni rango de incidencia): no es una fuente de incidencias.'
    return
  }
  if (!tieneDimension) {
    region.motivo = 'Sin columnas de días ni de fechas de incidencia.'
    return
  }

  // Una tabla por período con importes es un cuadro de remuneración.
  if ((meses > 0 || periodos > 0) && montos > 0) {
    region.motivo = 'Tabla por período con importes: es remuneración, no incidencias.'
    return
  }

  // Textos de remuneración en la columna de tipo: no son incidencias.
  if (montos > 0 && descripciones.length > 0) {
    const muestreo = muestrearTextos(region, descripciones[0], 20)
    const pagos = muestreo.filter((texto) => esConceptoRemunerativo(texto)).length
    if (muestreo.length > 0 && pagos / muestreo.length >= 0.6) {
      region.motivo = 'Los textos de la columna de tipo corresponden a conceptos de remuneración: no es una fuente de incidencias.'
      return
    }
  }

  // Sin tipo, solo se aceptan filas con un rango de fechas identificable.
  if (!tieneTipo && fechasSemanticas === 0) {
    region.motivo = 'Solo hay fechas sin semantica de inicio/fin y ningun tipo: se requiere revision manual.'
    return
  }

  // Una fila por trabajador con un único par inicio/fin describe al trabajador,
  // no una incidencia: son los datos de ingreso/cese del período.
  if (!tieneTipo) {
    const { filas, dnis } = muestrearFilas(region)
    if (filas > 0 && filas / Math.max(dnis, 1) <= 1.2) {
      region.motivo = 'Una fila por trabajador con fechas de inicio/fin: son datos del trabajador, no incidencias.'
      return
    }
  }

  region.aceptada = true
  region.motivo = 'Región de incidencias detectada por contenido.'
}

function muestrearFilas(region: Region, maximo: number = 200): { filas: number; dnis: number } {
  const documentos = new Set<string>()
  let filas = 0
  const hasta = Math.min(region.finBloque, region.filaDatos + maximo)
  for (let fila = region.filaDatos; fila <= hasta; fila++) {
    const dni = normalizarDocumento(region.ws.getRow(fila).getCell(region.colDocumento).value)
    if (!dni) continue
    filas++
    documentos.add(dni)
  }
  return { filas, dnis: documentos.size }
}

function muestrearTextos(region: Region, columna: PerfilColumna, maximo: number): string[] {
  const textos: string[] = []
  const hasta = Math.min(region.ws.rowCount || region.filaDatos, region.filaDatos + 200)
  for (let fila = region.filaDatos; fila <= hasta && textos.length < maximo; fila++) {
    const texto = normalizarTexto(region.ws.getRow(fila).getCell(columna.col).value)
    if (texto) textos.push(texto)
  }
  return textos
}

// ---------------------------------------------------------------------------
// Días laborables del rango (para validar columnas de días calculados)
// ---------------------------------------------------------------------------

function diasLaborables(desde: Date, hasta: Date, jornada: number, feriados: Set<string>): number {
  if (hasta < desde) return 0
  let total = 0
  const actual = new Date(desde)
  while (actual <= hasta) {
    const iso = actual.toISOString().slice(0, 10)
    const dow = actual.getUTCDay()
    const laborable = jornada === 6 ? dow !== 0 : jornada === 5 ? dow !== 0 && dow !== 6 : dow >= 1 && dow <= 4
    if (laborable && !feriados.has(iso)) total++
    actual.setUTCDate(actual.getUTCDate() + 1)
  }
  return total
}

// ---------------------------------------------------------------------------
// Lectura de registros
// ---------------------------------------------------------------------------

function columnaDocumento(region: Region): PerfilColumna | undefined {
  return region.columnas.find((columna) => columna.col === region.colDocumento)
}

function mejorColumna(region: Region, rol: Rol): PerfilColumna | undefined {
  const candidatas = columnasConRol(region, rol)
  if (candidatas.length === 0) return undefined
  return candidatas
    .slice()
    .sort((a, b) => b.filas - a.filas || a.col - b.col)[0]
}

function origenDe(region: Region, columna: PerfilColumna | undefined, fila: number): OrigenCampo | null {
  if (!columna) return null
  const valor = region.ws.getRow(fila).getCell(columna.col).value
  if (valorVacio(valor)) return null
  return {
    hoja: region.hoja,
    fila,
    columna: columna.col,
    encabezado: columna.encabezado,
    valor: valor instanceof Date ? fechaIso(valor) : valorATextoPlano(valor),
  }
}

function resolverFechas(region: Region): { inicio?: PerfilColumna; fin?: PerfilColumna; sinClasificar: PerfilColumna[] } {
  const inicio = mejorColumna(region, 'fechaInicio')
  const fin = mejorColumna(region, 'fechaFin')
  const sinClasificar = columnasConRol(region, 'fechaSinClasificar')
    .slice()
    .sort((a, b) => a.col - b.col)

  if (inicio || fin) return { inicio, fin, sinClasificar }

  if (sinClasificar.length === 1) return { inicio: sinClasificar[0], sinClasificar: [] }
  if (sinClasificar.length === 2) return { inicio: sinClasificar[0], fin: sinClasificar[1], sinClasificar: [] }
  return { sinClasificar }
}

function leerRegion(
  region: Region,
  opciones: Opciones,
  jornadas: Map<string, number>,
  feriadosPorDni: Map<string, Set<string>>,
): IncidenciaNormalizada[] {
  const loteId = opciones.loteId ?? 'lote-actual'
  const maxFilas = opciones.maxFilasPorRegion ?? 5000
  const maxPendientes = opciones.maxPendientesPorRegion ?? 200

  const colDocumento = columnaDocumento(region)
  if (!colDocumento) return []

  const colDescripcion = mejorColumna(region, 'descripcion')
  const colCodigo = mejorColumna(region, 'codigo')
  const colDias = mejorColumna(region, 'dias')
  const colDiasNeto = mejorColumna(region, 'diasNeto')
  const { inicio: colFechaInicio, fin: colFechaFin, sinClasificar } = resolverFechas(region)
  // El período se conserva tal como viene: columna "PERIODO" o una única
  // columna de mes. Varias columnas de mes son una matriz de días por mes, no
  // un rango de incidencia, y ya se descartan al evaluar la región.
  const colPeriodo = mejorColumna(region, 'periodo') ?? columnasConRol(region, 'mes')[0]

  const columnasEnterasSinRotular = sinClasificar.length === 0 && !colDiasNeto
    ? region.columnas.filter(
        (columna) =>
          region.rol.get(columna.col) === 'diasNeto' &&
          columna.filas > 0 &&
          fraccion(columna, 'enterosCorto') >= 0.95 &&
          fraccion(columna, 'montos') === 0
      )
    : []

  const registros: IncidenciaNormalizada[] = []
  let pendientes = 0
  const hasta = Math.min(region.finBloque, region.filaDatos + maxFilas)

  for (let fila = region.filaDatos; fila <= hasta; fila++) {
    const dni = normalizarDocumento(region.ws.getRow(fila).getCell(colDocumento.col).value)
    if (!dni) continue

    const descripcionOriginal = colDescripcion ? limpiarTextoValor(region.ws.getRow(fila).getCell(colDescripcion.col).value) : ''
    const descripcion = normalizarTexto(descripcionOriginal)
    const codigo = colCodigo ? valorATextoPlano(region.ws.getRow(fila).getCell(colCodigo.col).value).trim() : ''
    const diasCrudos = colDias ? numeroDe(region.ws.getRow(fila).getCell(colDias.col).value) : null
    const dias = diasCrudos !== null && Number.isFinite(diasCrudos) && diasCrudos > 0 ? diasCrudos : null
    const fechaInicio = colFechaInicio ? leerFecha(region.ws.getRow(fila).getCell(colFechaInicio.col).value) : null
    const fechaFin = colFechaFin ? leerFecha(region.ws.getRow(fila).getCell(colFechaFin.col).value) : null
    const periodo = colPeriodo ? leerPeriodo(region.ws.getRow(fila).getCell(colPeriodo.col).value) : null

    const origen: OrigenCampo[] = []
    const origenDocumento = origenDe(region, colDocumento, fila)
    if (origenDocumento) origen.push(origenDocumento)
    const origenDescripcion = origenDe(region, colDescripcion, fila)
    if (origenDescripcion) origen.push(origenDescripcion)
    const origenCodigo = origenDe(region, colCodigo, fila)
    if (origenCodigo) origen.push(origenCodigo)
    const origenDias = origenDe(region, colDias, fila)
    if (origenDias) origen.push(origenDias)
    const origenInicio = origenDe(region, colFechaInicio, fila)
    if (origenInicio) origen.push(origenInicio)
    const origenFin = origenDe(region, colFechaFin, fila)
    if (origenFin) origen.push(origenFin)
    if (periodo) {
      const origenPeriodo = origenDe(region, colPeriodo, fila)
      if (origenPeriodo) origen.push(origenPeriodo)
    }

    const textoRemunerativo = descripcion !== '' && esConceptoRemunerativo(descripcion)
    const hayRango = Boolean(fechaInicio && fechaFin && fechaFin >= fechaInicio)
    // Una fila sin tipo, sin días y sin rango solo se conserva si aporta un
    // período: queda como pendiente de revisión en lugar de desaparecer.
    const sinEvidencia = descripcion === '' && dias === null && !hayRango && !periodo
    if (sinEvidencia) continue
    // El tipo se conserva aunque falten los días: sigue siendo lo que dice el
    // archivo y es lo que hay que revisar.
    const faltaSoloCantidad = descripcion !== '' && dias === null && !hayRango

    let estado: EstadoIncidencia = 'identificada'
    let motivo = ''
    if (textoRemunerativo) {
      estado = 'no_identificada'
      motivo = 'El texto de la fila corresponde a un concepto de remuneración, no a una incidencia.'
    } else if (descripcion === '') {
      estado = 'no_identificada'
      motivo = 'La fila no indica el tipo de incidencia; se requiere revisión manual.'
    } else if (faltaSoloCantidad) {
      // El archivo identifica el tipo, pero no la cantidad de días ni un rango
      // de fechas: no se inventa el importe y queda pendiente de revisión.
      estado = 'no_identificada'
      motivo = periodo
        ? 'La fila no indica días ni rango de fechas: los días del período quedan pendientes de revisión.'
        : 'La fila no indica días, fechas ni período: los días quedan pendientes de revisión.'
    }

    let diasNeto: number | null = null
    if (colDiasNeto) {
      const valor = numeroDe(region.ws.getRow(fila).getCell(colDiasNeto.col).value)
      if (valor !== null && valor > 0) diasNeto = valor
    } else if (columnasEnterasSinRotular.length > 0) {
      diasNeto = elegirDiasCalculados(columnasEnterasSinRotular, fila, region, dias, fechaInicio, fechaFin, jornadas.get(dni), feriadosPorDni.get(dni) ?? new Set<string>())
    }
    if (diasNeto !== null) {
      const columna = columnasEnterasSinRotular.find((c) => numeroDe(region.ws.getRow(fila).getCell(c.col).value) === diasNeto)
      const origenNeto = origenDe(region, colDiasNeto ?? columna, fila)
      if (origenNeto) origen.push(origenNeto)
    }

    if (estado === 'no_identificada' && pendientes >= maxPendientes) continue
    if (estado === 'no_identificada') pendientes++

    registros.push({
      id: `${loteId}|${region.hoja}|${fila}|${dni}`,
      loteId,
      dni,
      tipo: estado === 'identificada' ? clasificarTipo(descripcion) : 'otro',
      descripcion: estado === 'identificada' || faltaSoloCantidad ? descripcionOriginal : '',
      codigo,
      dias,
      diasNeto,
      fechaInicio,
      fechaFin,
      periodo,
      estado,
      motivo,
      confianza: colDescripcion && colDescripcion.encabezado ? 'alta' : 'media',
      origen,
    })
  }

  return registros
}

function elegirDiasCalculados(
  candidatas: PerfilColumna[],
  fila: number,
  region: Region,
  dias: number | null,
  fechaInicio: Date | null,
  fechaFin: Date | null,
  jornada: number | undefined,
  feriados: Set<string>,
): number | null {
  const valores = candidatas
    .map((columna) => ({ columna, valor: numeroDe(region.ws.getRow(fila).getCell(columna.col).value) }))
    .filter((item): item is { columna: PerfilColumna; valor: number } => item.valor !== null && item.valor > 0)

  if (valores.length === 0) return null
  if (valores.length === 1) return valores[0].valor

  if (fechaInicio && fechaFin && fechaFin >= fechaInicio) {
    const laborables = diasLaborables(fechaInicio, fechaFin, jornada ?? 5, feriados)
    const coincidencia = valores.find((item) => Math.abs(item.valor - laborables) < 0.01)
    if (coincidencia) return coincidencia.valor
  }

  if (dias !== null) {
    const coincidencia = valores.find((item) => Math.abs(item.valor - dias) < 0.01)
    if (coincidencia) return coincidencia.valor
  }

  return valores[0].valor
}

// ---------------------------------------------------------------------------
// API
// ---------------------------------------------------------------------------


export function extraerIncidencias(workbook: ExcelJS.Workbook, opciones: Opciones = {}): ResultadoIncidencias {
  const regiones = construirRegiones(workbook)
  const aceptadas = regiones.filter((region) => region.aceptada)
  const jornadas = new Map<string, number>()
  const feriadosPorDni = new Map<string, Set<string>>()
  const registros: IncidenciaNormalizada[] = []
  const nombres = new Map<string, { apellidoPaterno: string; apellidoMaterno: string; nombres: string }>()

  for (const region of aceptadas) {
    for (const registro of leerRegion(region, opciones, jornadas, feriadosPorDni)) {
      registros.push(registro)
    }
  }

  // Los nombres del trabajador no son datos de incidencia: se toman de cualquier
  // tabla con documento, combinando el nombre completo con las columnas separadas.
  const crudos = new Map<string, NombresCrudos>()
  for (const region of regiones) {
    for (const [dni, parte] of leerNombres(region)) {
      const previo = crudos.get(dni)
      if (!previo) {
        crudos.set(dni, parte)
        continue
      }
      crudos.set(dni, {
        completo: previo.completo || parte.completo,
        split: previo.deIncidencia || !parte.deIncidencia ? previo.split ?? parte.split : parte.split,
        deIncidencia: previo.deIncidencia || parte.deIncidencia,
      })
    }
  }
  for (const [dni, parte] of crudos) {
    const { apP, apM, nom } = resolverNombres(parte)
    if (!apP && !apM && !nom) continue
    nombres.set(dni, { apellidoPaterno: apP, apellidoMaterno: apM, nombres: nom })
  }

  const porDni = new Map<string, IncidenciaNormalizada[]>()
  for (const registro of registros) {
    const lista = porDni.get(registro.dni) ?? []
    lista.push(registro)
    porDni.set(registro.dni, lista)
  }

  const resumenRegiones: ResumenRegion[] = aceptadas.map((region) => {
    const propios = registros.filter((registro) => registro.origen.some((o) => o.hoja === region.hoja && o.fila === registro.origen[0].fila && region.columnas.some((c) => c.col === o.columna)))
    const identificados = propios.filter((registro) => registro.estado === 'identificada').length
    return {
      hoja: region.hoja,
      filaEncabezado: region.filaEncabezado,
      columnas: region.columnas.map((columna) => `${columna.col}:${region.rol.get(columna.col)}:${columna.encabezado}`),
      registros: propios.length,
      identificados,
      pendientes: propios.length - identificados,
      motivo: region.motivo,
    }
  })

  return { registros, porDni, nombres, regiones: resumenRegiones }
}

type NombresSplit = { apP: string; apM: string; nom: string }
type NombresCrudos = { completo: string; split: NombresSplit | null; deIncidencia: boolean }

const PARTICULAS = new Set(['DE', 'DEL', 'LA', 'LAS', 'LOS', 'DAS', 'DOS', 'DI', 'DU', 'DA', 'DO', 'Y', 'VAN', 'VON', 'LE'])

const normalizarPalabras = (texto: string): string[] =>
  texto
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .split(/\s+/)
    .filter(Boolean)

const esSufijo = (completo: string[], trozo: string[]): boolean =>
  trozo.length > 0 && completo.length >= trozo.length && completo.slice(completo.length - trozo.length).join(' ') === trozo.join(' ')

// Las columnas separadas mandan cuando son coherentes con el nombre completo.
// Si no lo son (apellido materno que es una partícula, o nombres que repiten el
// apellido), se reparte el nombre completo usando el apellido paterno como ancla.
function resolverNombres(crudos: NombresCrudos): NombresSplit {
  const split = crudos.split
  if (!crudos.completo) return split ?? { apP: '', apM: '', nom: '' }

  const partes = crudos.completo.split(/\s+/).filter(Boolean)
  if (partes.length === 0) return split ?? { apP: '', apM: '', nom: '' }

  const partesNorm = normalizarPalabras(crudos.completo)
  const apP = split?.apP ?? ''
  const apPNorm = normalizarPalabras(apP)
  const apMNorm = normalizarPalabras(split?.apM ?? '')
  const nomNorm = normalizarPalabras(split?.nom ?? '')

  if (split && apPNorm.length > 0 && partesNorm.slice(0, apPNorm.length).join(' ') === apPNorm.join(' ')) {
    const resto = partesNorm.slice(apPNorm.length)
    const apMFiable = apMNorm.length > 0 && !PARTICULAS.has(apMNorm[0])
    if (apMFiable && resto.slice(0, apMNorm.length).join(' ') === apMNorm.join(' ')) {
      const cola = resto.slice(apMNorm.length)
      if (nomNorm.length === 0 || esSufijo(cola, nomNorm)) return split
    }
    // El apellido materno de la hoja no es fiable: se usa el nombre declarado
    // como ancla final y el resto queda como apellido.
    if (split.nom && esSufijo(resto, nomNorm) && resto.length > nomNorm.length) {
      return {
        apP,
        apM: resto.slice(0, resto.length - nomNorm.length).join(' '),
        nom: split.nom,
      }
    }
    if (resto.length >= 3) {
      return { apP, apM: resto.slice(0, resto.length - 2).join(' '), nom: resto.slice(resto.length - 2).join(' ') }
    }
    return { apP, apM: resto[0] ?? '', nom: resto.slice(1).join(' ') }
  }

  if (partes.length >= 3) {
    return { apP: partes[0], apM: partes[1], nom: partes.slice(2).join(' ') }
  }
  return { apP: partes[0], apM: '', nom: partes[1] ?? partes[0] }
}

function leerNombres(region: Region): Map<string, NombresCrudos> {
  const resultado = new Map<string, NombresCrudos>()
  const colDocumento = columnaDocumento(region)
  if (!colDocumento) return resultado

  const conTexto = (rol: Rol): PerfilColumna | undefined =>
    columnasConRol(region, rol).find((columna) => fraccion(columna, 'textos') >= 0.5)

  const colPaterno = conTexto('apellidoPaterno')
  const colMaterno = conTexto('apellidoMaterno')
  const colNombres = conTexto('nombres')
  const colCompleto = conTexto('nombre')
  if (!colPaterno && !colMaterno && !colNombres && !colCompleto) return resultado

  const deIncidencia = columnasConRol(region, 'descripcion').length > 0
  const hasta = Math.min(region.finBloque, region.filaDatos + 5000)
  for (let fila = region.filaDatos; fila <= hasta; fila++) {
    const dni = normalizarDocumento(region.ws.getRow(fila).getCell(colDocumento.col).value)
    if (!dni) continue

    const leer = (columna?: PerfilColumna): string =>
      columna ? limpiarTextoValor(region.ws.getRow(fila).getCell(columna.col).value) : ''

    const completoLeido = leer(colCompleto)
    const nombresLeidos = leer(colNombres)
    // Si el campo de nombres trae el nombre completo, no aporta nombres: se
    // descarta para no duplicar apellidos.
    const nombresEsCompleto =
      completoLeido !== '' && normalizarPalabras(nombresLeidos).join(' ') === normalizarPalabras(completoLeido).join(' ')

    const apP = leer(colPaterno)
    const apM = leer(colMaterno)
    const nom = nombresEsCompleto ? '' : nombresLeidos
    const parte: NombresCrudos = {
      completo: completoLeido || (nombresEsCompleto ? nombresLeidos : ''),
      split: apP || apM || nom ? { apP, apM, nom } : null,
      deIncidencia,
    }
    if (!parte.completo && !parte.split) continue

    const previo = resultado.get(dni)
    if (!previo) {
      resultado.set(dni, parte)
      continue
    }
    resultado.set(dni, {
      completo: previo.completo || parte.completo,
      split: previo.deIncidencia || !parte.deIncidencia ? previo.split ?? parte.split : parte.split,
      deIncidencia: previo.deIncidencia || parte.deIncidencia,
    })
  }
  return resultado
}
