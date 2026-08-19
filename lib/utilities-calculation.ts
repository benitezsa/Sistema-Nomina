import type { Employee, EmployeeUtilityResult, UtilityParameters } from '@/components/utilities/types'

export const money = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN', minimumFractionDigits: 2 })
export const number = new Intl.NumberFormat('es-PE')
export function calculateFund(p: UtilityParameters) { return p.income * (p.legalPercent / 100) }
export function calculateResults(employees: Employee[], p: UtilityParameters): EmployeeUtilityResult[] {
  const fund = calculateFund(p)
  const half = fund / 2
  const totalDays = employees.reduce((sum, employee) => sum + employee.days, 0)
  const totalRemuneration = employees.reduce((sum, employee) => sum + employee.remuneration, 0)
  return employees.map((employee) => {
    const daysFactor = employee.days / totalDays
    const remunerationFactor = employee.remuneration / totalRemuneration
    const preliminary = half * (daysFactor + remunerationFactor)
    const cap = employee.remuneration * p.capMonths
    const finalAmount = Math.min(preliminary, cap)
    return { ...employee, daysFactor, remunerationFactor, preliminary, cap, finalAmount, remainder: Math.max(0, preliminary - finalAmount), capApplied: preliminary > cap }
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
  return { fund, distributed, remainder: fund - distributed, capped: results.filter((result) => result.capApplied).length }
}
