export type ExerciseStatus = 'Borrador' | 'En cálculo' | 'Calculado' | 'Cerrado'
export type EmployeeStatus = 'Completo' | 'Pendiente' | 'Observado'

export type Company = { id: string; name: string; ruc: string; activity: string; address: string }
export type CompanyStatus = 'Activo' | 'Pendiente' | 'Inactivo'
export type CompanyClient = { id: string; legalName: string; tradeName: string; ruc: string; activity: string; contact: string; role: string; phone: string; email: string; address: string; status: CompanyStatus; employees: number; lastExercise: number }
export type Exercise = { id: string; year: number; company: Company; status: ExerciseStatus; createdAt: string; updatedAt: string; employeeCount: number; fund: number }
export type Employee = { id: string; code: string; name: string; role: string; department: string; days: number; remuneration: number; status: EmployeeStatus; capApplied?: boolean }
export type UtilityParameters = { income: number; legalPercent: number; averageEmployees: number; daysBase: number; remunerationMonths: number; capMonths: number }
export type EmployeeUtilityResult = Employee & { daysFactor: number; remunerationFactor: number; preliminary: number; cap: number; finalAmount: number; remainder: number }
export type Remainder = { employeeId: string; employeeName: string; preliminary: number; cap: number; excess: number; reason: string }
export type AuditLog = { id: string; date: string; user: string; action: string; detail: string; status: 'Completado' | 'Pendiente' }
export type CalculationVersion = { id: string; version: string; date: string; user: string; reason: string; total: number }
export type ViewKey = 'Empresas' | 'Configuración' | 'Trabajadores' | 'Cálculo' | 'Resultados' | 'Remanentes' | 'Reportes' | 'Auditoría'
export type ToastMessage = { title: string; description: string }
export const VIEW_KEYS: ViewKey[] = ['Empresas', 'Configuración', 'Trabajadores', 'Cálculo', 'Resultados', 'Remanentes', 'Reportes', 'Auditoría']
