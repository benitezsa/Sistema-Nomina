import type * as ExcelJS from 'exceljs'
import type { Employee } from '@/components/utilities/types'
import {
  consolidarTrabajadores,
  extraerDias,
  extraerRemuneraciones,
  inferirAnioEjercicio,
  type ImportedWorker,
} from '@/lib/import-utilidades'
import type { OpcionesRemuneracionAtomica } from '@/lib/import-utilidades-atomico'
import { configuracion, type ConfiguracionCalculo } from './config'
import type { TrabajadorNormalizado } from './modelo'
import { normalizarTrabajadores, type OverridesNormalizacion } from './normalizar'
import { aEmployee, calcularTrabajadores, type TrabajadorCalculado } from './trabajador'

export type ResultadoExtraccion = {
  remuneraciones: ImportedWorker[]
  dias: ImportedWorker[]
  consolidado: ImportedWorker[]
}

/** Extrae y consolida (sin calcular todavía): produce datos base normalizables. */
export function extraerYConsolidar(
  workbook: ExcelJS.Workbook,
  actuales: ImportedWorker[] = [],
  opciones: OpcionesRemuneracionAtomica = {},
): ResultadoExtraccion {
  const remuneraciones = extraerRemuneraciones(workbook, actuales, opciones).trabajadores
  const dias = extraerDias(workbook, remuneraciones).trabajadores
  const consolidado = consolidarTrabajadores(dias, remuneraciones)
  return { remuneraciones, dias, consolidado }
}

export type ResultadoCalculo = {
  normalizados: TrabajadorNormalizado[]
  calculados: TrabajadorCalculado[]
  empleados: Employee[]
}

/**
 * Normaliza los trabajadores consolidados y ejecuta el motor de cálculo.
 * Es la única puerta por la que el modelo normalizado llega a los resultados.
 */
export function calcularEmpleados(
  consolidado: ImportedWorker[],
  config: ConfiguracionCalculo,
  overridesPorDni: Record<string, OverridesNormalizacion> = {},
): ResultadoCalculo {
  const normalizados = normalizarTrabajadores(consolidado, config, overridesPorDni)
  const calculados = calcularTrabajadores(normalizados, config)
  return {
    normalizados,
    calculados,
    empleados: calculados.map(aEmployee),
  }
}

/**
 * Pipeline completo: Excel → extracción → normalización → cálculo → empleados.
 * El motor de resultados solo ve el modelo normalizado, nunca el Excel.
 */
export function procesarUtilidades(
  workbook: ExcelJS.Workbook,
  actuales: ImportedWorker[] = [],
  opciones: OpcionesRemuneracionAtomica = {},
  configOverrides: Partial<ConfiguracionCalculo> = {},
): ResultadoExtraccion & ResultadoCalculo {
  const extraccion = extraerYConsolidar(workbook, actuales, opciones)
  const anioEjercicio = configOverrides.anioEjercicio ?? inferirAnioEjercicio(workbook)
  const config = configuracion({ ...configOverrides, anioEjercicio })
  return { ...extraccion, ...calcularEmpleados(extraccion.consolidado, config) }
}

export type { TrabajadorCalculado, TrabajadorNormalizado }
