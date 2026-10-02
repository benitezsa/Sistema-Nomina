import {
  calcularDiasEfectivos,
  calcularDiasLaborablesRango,
  calcularDiasNoLaborados,
} from '@/lib/utilities-calculation'
import { resolverFeriados, type ConfiguracionCalculo } from './config'
import type { TrabajadorNormalizado } from './modelo'

export type OrigenDias = 'base-mensual' | 'calendario' | 'sin-datos'

export type ResultadoDias = {
  /** Días laborables/pósibles del período. */
  laborables: number
  /** Días no laborados por incidencias. */
  noLaborados: number
  /** Días efectivos = laborables − no laborados. */
  efectivos: number
  /** De dónde salió el cálculo, para trazabilidad (no para el resultado). */
  origen: OrigenDias
}

function aFecha(iso: string | null): Date | null {
  if (!iso) return null
  const fecha = new Date(`${iso}T00:00:00.000Z`)
  return Number.isNaN(fecha.getTime()) ? null : fecha
}

function rangoDe(t: TrabajadorNormalizado): { inicio: Date; fin: Date } | null {
  const ingreso = aFecha(t.fechaIngreso)
  const cese = aFecha(t.fechaCese)
  // Sin ninguna fecha real no se puede reconstruir el calendario: se usa el
  // calendario laborable que aporte el archivo (o cero).
  if (!ingreso && !cese) return null
  const anio = ingreso?.getUTCFullYear() ?? cese?.getUTCFullYear() ?? t.anioEjercicio ?? null
  if (anio === null) return null
  const inicio = ingreso ?? new Date(Date.UTC(anio, 0, 1))
  const fin = cese ?? new Date(Date.UTC(anio, 11, 31))
  if (fin < inicio) return null
  return { inicio, fin }
}

/**
 * Calcula los días a partir de datos base:
 *  - Si el archivo aporta días laborables por mes, se suman (nunca se lee un
 *    "total días"): esos días son el calendario laborable base.
 *  - Si no, y hay fecha de ingreso/cese + jornada, se cuenta el calendario
 *    laborable del rango (jornada semanal y feriados).
 *  - Los días no laborados salen de las incidencias.
 *  - Días efectivos = laborables − no laborados.
 */
export function calcularDias(t: TrabajadorNormalizado, config: ConfiguracionCalculo): ResultadoDias {
  const feriados = new Set<string>([...resolverFeriados(config), ...t.feriados])
  const aplicables = t.incidencias.filter((inc) => inc.aplicable)
  const noLaborados = calcularDiasNoLaborados(
    aplicables.map((inc) => ({
      cantidadDias: inc.cantidadDias,
      diasNeto: inc.diasNeto ?? undefined,
      fechaInicio: inc.fechaInicio,
      fechaFin: inc.fechaFin,
    })),
    t.jornada,
    feriados,
  )

  let laborables = 0
  let origen: OrigenDias = 'sin-datos'

  // Prioridad: el calendario laborable se recalcula desde las fechas de ingreso
  // y cese + jornada + feriados (dato base). Los días por mes del archivo solo
  // se usan si no hay rango de fechas que permita recalcular.
  const rango = rangoDe(t)
  if (rango) {
    if (t.jornadaPeriodos.length > 0) {
      laborables = t.jornadaPeriodos.reduce((suma, periodo) => {
        const desde = aFecha(periodo.desde)
        const hasta = aFecha(periodo.hasta)
        if (!desde || !hasta) return suma
        const inicio = desde > rango.inicio ? desde : rango.inicio
        const fin = hasta < rango.fin ? hasta : rango.fin
        if (fin < inicio) return suma
        return suma + calcularDiasLaborablesRango(inicio, fin, periodo.jornada, feriados)
      }, 0)
    } else {
      laborables = calcularDiasLaborablesRango(rango.inicio, rango.fin, t.jornada, feriados)
    }
    origen = 'calendario'
  } else if (t.diasBase.length > 0) {
    laborables = t.diasBase.reduce((suma, dia) => suma + dia.dias, 0)
    origen = 'base-mensual'
  }

  return {
    laborables,
    noLaborados,
    efectivos: calcularDiasEfectivos(laborables, noLaborados),
    origen,
  }
}
