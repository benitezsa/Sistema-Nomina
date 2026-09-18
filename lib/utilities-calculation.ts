import type { Employee, EmployeeUtilityResult, QuintaCategoriaResult, QuintaCategoriaTramo, UtilityParameters } from '@/components/utilities/types'

export const money = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN', minimumFractionDigits: 2 })
export const number = new Intl.NumberFormat('es-PE')

export const TRAMOS_QUINTA_CATEGORIA: QuintaCategoriaTramo[] = [
  { desde: 0, hasta: 5, tasa: 0.08 },
  { desde: 5, hasta: 20, tasa: 0.14 },
  { desde: 20, hasta: 35, tasa: 0.17 },
  { desde: 35, hasta: 45, tasa: 0.20 },
  { desde: 45, hasta: null, tasa: 0.30 },
]

export const UIT_DEFAULT = 5350
export const UIT_DEDUCCION = 7

export function calcularBaseImponible(remuneracionAnual: number, uit: number): number {
  return Math.max(0, remuneracionAnual - uit * UIT_DEDUCCION)
}

export function calcularTasaMarginal(baseImponible: number, uit: number): number {
  let tasa = 0
  for (const tramo of TRAMOS_QUINTA_CATEGORIA) {
    if (tramo.hasta === null) {
      if (baseImponible >= tramo.desde * uit) tasa = tramo.tasa
      break
    }
    if (baseImponible > tramo.desde * uit) {
      tasa = tramo.tasa
    }
    if (baseImponible <= tramo.hasta * uit) break
  }
  return tasa
}

export function calcularQuintaCategoria(utilidadBruta: number, remuneracionAnual: number, uit: number = UIT_DEFAULT): QuintaCategoriaResult {
  const deduccion7UIT = uit * UIT_DEDUCCION
  const baseImponible = calcularBaseImponible(remuneracionAnual, uit)
  const tasaMarginal = calcularTasaMarginal(baseImponible, uit)
  const impuesto = utilidadBruta * tasaMarginal

  const tramosAplicados: QuintaCategoriaResult['tramosAplicados'] = [{
    tramo: TRAMOS_QUINTA_CATEGORIA.find((t) => t.tasa === tasaMarginal) ?? TRAMOS_QUINTA_CATEGORIA[0],
    montoTramo: utilidadBruta,
    impuestoTramo: impuesto,
  }]

  return {
    remuneracionAnual,
    deduccion7UIT,
    baseImponible,
    tasaMarginal,
    impuesto: Math.round(impuesto * 100) / 100,
    tramosAplicados,
  }
}

export function calcularDiasLaborables(diasTrabajadosTotal: number, diasNoLaborados: number): number {
  return diasTrabajadosTotal + diasNoLaborados
}

export function calcularDiasNoLaboradosIncidencias(incidencias: { cantidadDias: number }[]): number {
  return incidencias.reduce((total, inc) => total + inc.cantidadDias, 0)
}

// Días de la semana según la jornada semanal (6 | 5 | 4).
// Counting in UTC to avoid timezone shifts.
function esDiaLaborable(fecha: Date, jornada: number): boolean {
  const dow = fecha.getUTCDay()
  if (jornada === 6) return dow !== 0
  if (jornada === 5) return dow !== 0 && dow !== 6
  return dow >= 1 && dow <= 4
}

// Cuenta los días laborables reales dentro de un rango [desde, hasta] inclusive,
// teniendo en cuenta la jornada semanal y los feriados (fechas 'YYYY-MM-DD').
export function calcularDiasLaborablesRango(
  desde: Date,
  hasta: Date,
  jornada: number = 5,
  feriados: ReadonlySet<string> = new Set()
): number {
  if (!desde || !hasta || hasta < desde) return 0
  const inicio = new Date(Date.UTC(desde.getUTCFullYear(), desde.getUTCMonth(), desde.getUTCDate()))
  const fin = new Date(Date.UTC(hasta.getUTCFullYear(), hasta.getUTCMonth(), hasta.getUTCDate()))
  let total = 0
  const cur = new Date(inicio)
  while (cur <= fin) {
    const iso = cur.toISOString().slice(0, 10)
    if (esDiaLaborable(cur, jornada) && !feriados.has(iso)) total++
    cur.setUTCDate(cur.getUTCDate() + 1)
  }
  return total
}

// Suma los días no laborados de cada incidencia.
// Prioridad por incidencia:
//  1. diasNeto: valor de días no laborados ya calculado por el archivo
//     (columna N del formato CITIKOLD / columna laborables del reporte PLAME).
//  2. Recalculo del tramo por fechas según jornada y feriados.
//  3. cantidadDias como último recurso.
// Esto evita descuentos inventados o dobles: cada concepto se descuenta una sola vez.
export function calcularDiasNoLaborados(
  incidencias: { cantidadDias: number; fechaInicio?: Date | null; fechaFin?: Date | null; diasNeto?: number }[],
  jornada: number = 5,
  feriados: ReadonlySet<string> = new Set()
): number {
  return incidencias.reduce((total, inc) => {
    if (inc.diasNeto !== undefined && inc.diasNeto !== null) {
      return total + inc.diasNeto
    }
    if (inc.fechaInicio && inc.fechaFin) {
      return total + calcularDiasLaborablesRango(inc.fechaInicio, inc.fechaFin, jornada, feriados)
    }
    return total + (inc.cantidadDias ?? 0)
  }, 0)
}

export function calcularDiasEfectivos(diasTrabajadosTotal: number, diasNoLaborados: number): number {
  return Math.max(0, diasTrabajadosTotal - diasNoLaborados)
}

export function calcularRemuneracionMensual(remuneracionAnual: number, mesesPresentes: number): number {
  if (mesesPresentes === 0) return 0
  return remuneracionAnual / mesesPresentes
}

export function calcularRemuneracionComputable(remuneracionMensual: number, mesesPresentes: number): number {
  return remuneracionMensual * mesesPresentes
}

export function calcularJornadaPeriodos(periodos: { desde: string; hasta: string; jornada: 6 | 5 | 4 }[], year: number): { desde: string; hasta: string; jornada: 6 | 5 | 4; meses: number }[] {
  return periodos.map(p => {
    const desde = new Date(p.desde)
    const hasta = new Date(p.hasta)
    const meses = Math.max(1, Math.round((hasta.getTime() - desde.getTime()) / (30.44 * 86400000)) + 1)
    return { ...p, meses: Math.min(meses, 12) }
  })
}

export function calculateFund(p: UtilityParameters) { return p.income * (p.legalPercent / 100) }

export function calculateResults(employees: Employee[], p: UtilityParameters): EmployeeUtilityResult[] {
  const fund = calculateFund(p)
  const half = fund / 2

  const totalDiasEfectivos = employees.reduce((sum, e) => sum + (e.diasEfectivos ?? e.days), 0)
  const totalRemuneracion = employees.reduce((sum, e) => sum + e.remuneration, 0)

  if (totalDiasEfectivos === 0 || totalRemuneracion === 0) {
    return employees.map(e => ({
      ...e,
      daysFactor: 0,
      remunerationFactor: 0,
      preliminary: 0,
      cap: e.remuneration * p.capMonths,
      finalAmount: 0,
      remainder: 0,
      diasEfectivos: e.diasEfectivos ?? e.days,
      remuneracionComputable: e.remuneration,
      distribucion: {
        utilidadPorDias: 0,
        utilidadPorRemuneraciones: 0,
        utilidadBruta: 0,
        quintaCategoria: 0,
        utilidadNeta: 0,
      },
    }))
  }

  return employees.map((employee) => {
    const diasEfectivos = employee.diasEfectivos ?? employee.days
    const remuneracionComputable = employee.remuneration

    const daysFactor = diasEfectivos / totalDiasEfectivos
    const remunerationFactor = remuneracionComputable / totalRemuneracion

    const utilidadPorDias = half * daysFactor
    const utilidadPorRemuneraciones = half * remunerationFactor
    const utilidadBruta = utilidadPorDias + utilidadPorRemuneraciones

    const cap = employee.remuneration * p.capMonths
    const preliminar = utilidadBruta
    const finalAmount = Math.min(preliminar, cap)
    const remainder = Math.max(0, preliminar - finalAmount)

    const quintaCategoria = calcularQuintaCategoria(finalAmount, remuneracionComputable)
    const utilidadNeta = Math.max(0, finalAmount - quintaCategoria.impuesto)

    return {
      ...employee,
      daysFactor,
      remunerationFactor,
      preliminary: preliminar,
      cap,
      finalAmount: Math.round(finalAmount * 100) / 100,
      remainder: Math.round(remainder * 100) / 100,
      capApplied: preliminar > cap,
      diasEfectivos,
      remuneracionComputable,
      distribucion: {
        utilidadPorDias: Math.round(utilidadPorDias * 100) / 100,
        utilidadPorRemuneraciones: Math.round(utilidadPorRemuneraciones * 100) / 100,
        utilidadBruta: Math.round(utilidadBruta * 100) / 100,
        quintaCategoria: quintaCategoria.impuesto,
        utilidadNeta: Math.round(utilidadNeta * 100) / 100,
      },
      quintaCategoria,
    }
  })
}

export function validateParameters(p: UtilityParameters, employeeCount: number) {
  const errors: string[] = []
  if (p.income <= 0) errors.push('La renta neta debe ser mayor a cero.')
  if (p.averageEmployees <= 20) errors.push('El promedio de trabajadores debe ser mayor a 20.')
  if (employeeCount === 0) errors.push('Debe existir al menos un trabajador.')
  return errors
}

export function getTotals(results: EmployeeUtilityResult[], p: UtilityParameters) {
  const fund = calculateFund(p)
  const distributed = results.reduce((sum, result) => sum + result.finalAmount, 0)
  const totalQuintaCategoria = results.reduce((sum, result) => sum + (result.quintaCategoria?.impuesto ?? 0), 0)
  return {
    fund,
    distributed,
    remainder: fund - distributed,
    capped: results.filter((result) => result.capApplied).length,
    totalQuintaCategoria: Math.round(totalQuintaCategoria * 100) / 100,
  }
}
