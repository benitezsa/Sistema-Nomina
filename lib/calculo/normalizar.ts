import { months, type ImportedWorker, type IncidenciaImportada } from '@/lib/import-utilidades'
import { normalizarJornada, type ConfiguracionCalculo } from './config'
import type {
  ConceptoRemunerativo,
  DiasBaseMes,
  IncidenciaBase,
  TrabajadorNormalizado,
} from './modelo'

/** Overrides que la interfaz permite editar antes de recalcular. */
export type OverridesNormalizacion = {
  jornada?: number
  fechaInicio?: string
  fechaCese?: string
  /** Meses trabajados seleccionados en la UI (si se editaron). */
  meses?: string[]
  /** Incidencias editadas en la UI (reemplazan a las importadas). */
  incidencias?: IncidenciaImportada[]
}

function aFechaIso(valor: unknown): string | null {
  if (!valor) return null
  if (valor instanceof Date) {
    if (Number.isNaN(valor.getTime())) return null
    return valor.toISOString().slice(0, 10)
  }
  if (typeof valor === 'number') {
    if (valor <= 20000 || valor >= 80000) return null
    return new Date(Date.UTC(1899, 11, 30) + valor * 86_400_000).toISOString().slice(0, 10)
  }
  if (typeof valor === 'string') {
    const texto = valor.trim()
    if (!texto) return null
    const iso = /^(\d{4})-(\d{2})-(\d{2})/.exec(texto)
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`
    const fecha = new Date(texto)
    return Number.isNaN(fecha.getTime()) ? null : fecha.toISOString().slice(0, 10)
  }
  if (typeof valor === 'object' && valor !== null) {
    const celda = valor as { result?: unknown; text?: unknown }
    if ('result' in valor) return aFechaIso(celda.result)
    if ('text' in valor) return aFechaIso(celda.text)
  }
  return null
}

function fechaDesdeIncidencia(valor: unknown): Date | null {
  if (!valor) return null
  if (valor instanceof Date) return Number.isNaN(valor.getTime()) ? null : valor
  const iso = aFechaIso(valor)
  return iso ? new Date(`${iso}T00:00:00.000Z`) : null
}

function normalizarIncidencia(inc: IncidenciaImportada): IncidenciaBase {
  return {
    descripcion: inc.descripcion,
    cantidadDias: inc.cantidadDias,
    diasNeto: inc.diasNeto ?? null,
    fechaInicio: fechaDesdeIncidencia(inc.fechaInicio),
    fechaFin: fechaDesdeIncidencia(inc.fechaFin),
    mes: inc.periodo?.mes ?? null,
    anio: inc.periodo?.anio ?? null,
    aplicable: inc.estado !== 'no_identificada',
  }
}

/**
 * Convierte un trabajador extraído de cualquier Excel al modelo normalizado.
 * Los totales que traía el archivo pasan a `referencias`; nunca a las entradas
 * del cálculo.
 */
export function normalizarTrabajador(
  worker: ImportedWorker,
  config: ConfiguracionCalculo,
  overrides: OverridesNormalizacion = {},
): TrabajadorNormalizado {
  const jornada = normalizarJornada(overrides.jornada ?? worker.jornada, config.jornadaPorDefecto)
  const anio = config.anioEjercicio

  const conceptos: ConceptoRemunerativo[] = months
    .map((mes, indice) => ({
      concepto: 'Remuneración',
      mes: indice + 1,
      anio,
      monto: worker.remuneraciones?.[mes] ?? 0,
      naturaleza: 'remunerativo' as const,
    }))
    .filter((concepto) => concepto.monto !== 0)

  // Si el archivo solo declara un total anual (sin desglose mensual), ese total
  // es el dato base disponible: se conserva como un único concepto sin mes.
  if (conceptos.length === 0 && (worker.remuneraciones?.total ?? 0) !== 0) {
    conceptos.push({
      concepto: 'Remuneración declarada',
      mes: null,
      anio,
      monto: worker.remuneraciones.total,
      naturaleza: 'remunerativo',
    })
  }

  const mesesPermitidos = overrides.meses ? new Set(overrides.meses) : null
  // Días laborables base del mes. Se prioriza `laborablesPorMes` (calendario
  // laborable aportado por el archivo); `diasTrabajados` es un resultado y solo
  // se usa como último recurso cuando no hay calendario ni fechas.
  const diasBase: DiasBaseMes[] = months
    .map((mes, indice) => ({
      mes: indice + 1,
      anio,
      dias: worker.laborablesPorMes?.[mes] ?? worker.diasTrabajados?.[mes] ?? 0,
    }))
    .filter((dia) => dia.dias > 0 && (!mesesPermitidos || mesesPermitidos.has(months[dia.mes - 1])))

  const incidenciasFuente = overrides.incidencias ?? worker.incidencias ?? []

  return {
    dni: worker.dni,
    apellidoPaterno: worker.apellidoPaterno,
    apellidoMaterno: worker.apellidoMaterno,
    nombres: worker.nombres,
    empresa: '',
    fechaIngreso: overrides.fechaInicio ?? aFechaIso(worker.fechaInicio),
    fechaCese: overrides.fechaCese ?? aFechaIso(worker.fechaCese),
    jornada,
    jornadaPeriodos: [],
    anioEjercicio: anio,
    feriados: worker.feriados ?? [],
    incidencias: incidenciasFuente.map(normalizarIncidencia),
    conceptos,
    diasBase,
    referencias: {
      ...(typeof worker.diasPosibles === 'number' ? { diasPosibles: worker.diasPosibles } : {}),
      ...(typeof worker.diasEfectivosReferencia === 'number' ? { diasEfectivos: worker.diasEfectivosReferencia } : {}),
      ...(typeof worker.diasNoLaboradosReferencia === 'number' ? { diasNoLaborados: worker.diasNoLaboradosReferencia } : {}),
      ...(worker.remuneraciones?.total ? { remuneracionTotal: worker.remuneraciones.total } : {}),
    },
  }
}

export function normalizarTrabajadores(
  workers: ImportedWorker[],
  config: ConfiguracionCalculo,
  overridesPorDni: Record<string, OverridesNormalizacion> = {},
): TrabajadorNormalizado[] {
  return workers.map((worker) => normalizarTrabajador(worker, config, overridesPorDni[worker.dni] ?? {}))
}
