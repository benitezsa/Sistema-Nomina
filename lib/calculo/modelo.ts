import type { Jornada, JornadaPeriodo } from '@/components/utilities/types'

// ---------------------------------------------------------------------------
// MODELO NORMALIZADO
//
// Es la frontera entre la extracción (que sí conoce hojas y columnas) y el
// motor de cálculo (que no debe conocerlas). Cualquier Excel, sea GR, CITIKOLD
// u otro, debe terminar produciendo este modelo.
//
// Regla clave: los datos que son RESULTADO calculado por el archivo (total de
// días efectivos, total de remuneración) no entran al motor; solo se conservan
// como "referencias" para auditoría/validación.
// ---------------------------------------------------------------------------

/** Naturaleza de un concepto remunerativo según el catálogo del importador. */
export type NaturalezaConcepto = 'remunerativo' | 'no_remunerativo' | 'desconocido'

/** Monto base de un concepto, con su período. */
export type ConceptoRemunerativo = {
  concepto: string
  /** Mes 1..12; null si el archivo no lo aporta. */
  mes: number | null
  anio: number | null
  monto: number
  naturaleza: NaturalezaConcepto
}

/** Días laborables/pósibles que el archivo aporta para un mes. Es dato base. */
export type DiasBaseMes = {
  mes: number
  anio: number | null
  dias: number
}

/** Incidencia base (falta, vacaciones, licencia, etc.). */
export type IncidenciaBase = {
  descripcion: string
  cantidadDias: number
  diasNeto: number | null
  fechaInicio: Date | null
  fechaFin: Date | null
  mes: number | null
  anio: number | null
  /** false cuando la incidencia quedó sin identificar: no descuenta días. */
  aplicable: boolean
}

/**
 * Valores que el archivo entrega ya calculados. NO se usan como entrada del
 * cálculo; solo sirven para comparar contra lo que el sistema calcula.
 */
export type ReferenciasImportadas = {
  diasPosibles?: number
  diasEfectivos?: number
  diasNoLaborados?: number
  remuneracionTotal?: number
}

/**
 * Trabajador normalizado: todo dato de entrada que el motor necesita, sin
 * ninguna dependencia de hojas, columnas ni años embebidos.
 */
export type TrabajadorNormalizado = {
  dni: string
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string
  empresa: string

  fechaIngreso: string | null
  fechaCese: string | null
  jornada: Jornada
  jornadaPeriodos: JornadaPeriodo[]

  anioEjercicio: number | null
  feriados: string[]

  incidencias: IncidenciaBase[]

  /** Base remunerativa: conceptos/montos por período. */
  conceptos: ConceptoRemunerativo[]

  /** Base de asistencia: días laborables/pósibles por mes aportados por el archivo. */
  diasBase: DiasBaseMes[]

  referencias: ReferenciasImportadas
}
