import type { Jornada, QuintaCategoriaTramo } from '@/components/utilities/types'

// ---------------------------------------------------------------------------
// Configuración del cálculo
//
// Todo valor legal (UIT, tramos de quinta categoría), el año del ejercicio y
// los feriados viven aquí. El motor de cálculo recibe esta configuración; no
// conoce hojas, columnas ni años embebidos en el código.
// ---------------------------------------------------------------------------

export const UIT_DEFAULT = 5350
export const UIT_DEDUCCION = 7

export const TRAMOS_QUINTA_CATEGORIA: QuintaCategoriaTramo[] = [
  { desde: 0, hasta: 5, tasa: 0.08 },
  { desde: 5, hasta: 20, tasa: 0.14 },
  { desde: 20, hasta: 35, tasa: 0.17 },
  { desde: 35, hasta: 45, tasa: 0.20 },
  { desde: 45, hasta: null, tasa: 0.30 },
]

export type FuenteFeriados = readonly string[] | ((anio: number | null) => readonly string[])

export type ConfiguracionCalculo = {
  /** Año del ejercicio. Puede ser null si el archivo no lo aporta. */
  anioEjercicio: number | null
  /** UIT vigente del ejercicio. */
  uit: number
  /** Feriados en formato ISO 'YYYY-MM-DD'. */
  feriados: FuenteFeriados
  /** Jornada semanal por defecto (6 | 5 | 4) cuando el archivo no la aporta. */
  jornadaPorDefecto: Jornada
}

export const CONFIGURACION_POR_DEFECTO: ConfiguracionCalculo = {
  anioEjercicio: null,
  uit: UIT_DEFAULT,
  feriados: [],
  jornadaPorDefecto: 5,
}

export function resolverFeriados(config: ConfiguracionCalculo): Set<string> {
  const fuente = config.feriados
  const lista = typeof fuente === 'function' ? fuente(config.anioEjercicio) : fuente
  return new Set(lista.map((fecha) => fecha.slice(0, 10)))
}

export function configuracion(overrides: Partial<ConfiguracionCalculo> = {}): ConfiguracionCalculo {
  return { ...CONFIGURACION_POR_DEFECTO, ...overrides }
}

/** Normaliza una jornada leída de un archivo (5 | 6 | 4) o de la UI. */
export function normalizarJornada(valor: unknown, porDefecto: Jornada = 5): Jornada {
  if (valor === 6 || valor === 4 || valor === 5) return valor
  if (typeof valor === 'string') {
    const texto = valor.trim().toUpperCase()
    if (texto === '6' || texto.includes('S')) return 6
    if (texto === '4') return 4
    if (texto === '5') return 5
  }
  if (typeof valor === 'number' && (valor === 6 || valor === 5 || valor === 4)) return valor
  return porDefecto
}
