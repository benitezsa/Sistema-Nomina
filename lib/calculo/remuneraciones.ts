import { calcularRemuneracionMensual } from '@/lib/utilities-calculation'
import type { TrabajadorNormalizado } from './modelo'

export type ResultadoRemuneracion = {
  /** Remuneración computable anual (suma de conceptos remunerativos). */
  anual: number
  /** Remuneración mensual promedio (anual / meses con remuneración). */
  mensual: number
  /** Meses con remuneración computable. */
  mesesConRemuneracion: number
  /** Detalle por mes, para trazabilidad. */
  porMes: { mes: number | null; anio: number | null; monto: number }[]
}

/**
 * Determina la remuneración computable a partir de los conceptos base.
 *
 * Reglas:
 *  - Solo suman los conceptos de naturaleza remunerativa (o desconocida, que el
 *    importador ya clasificó como computable al construir el monto mensual).
 *  - Se excluyen los conceptos no remunerativos.
 *  - La remuneración mensual se deriva del anual entre los meses con monto
 *    (función `calcularRemuneracionMensual`), no se lee de una columna TOTAL.
 */
export function calcularRemuneracionComputable(t: TrabajadorNormalizado): ResultadoRemuneracion {
  const computables = t.conceptos.filter(
    (concepto) => concepto.naturaleza !== 'no_remunerativo' && concepto.monto !== 0,
  )

  const porMes: ResultadoRemuneracion['porMes'] = []
  const meses = new Set<number>()
  let anual = 0

  for (const concepto of computables) {
    anual += concepto.monto
    if (concepto.mes !== null) meses.add(concepto.mes)
    porMes.push({ mes: concepto.mes, anio: concepto.anio, monto: concepto.monto })
  }

  const mesesConRemuneracion = meses.size
  const mesesBase = mesesConRemuneracion > 0 ? mesesConRemuneracion : 12
  const mensual = calcularRemuneracionMensual(anual, mesesBase)

  return { anual, mensual, mesesConRemuneracion, porMes }
}
