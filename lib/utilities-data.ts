import type { AuditLog, CalculationVersion, Company, CompanyClient, Employee, Exercise, UtilityParameters } from '@/components/utilities/types'

// Datos reales pendientes de conexión. No se muestran datos ficticios.
export const company: Company = { id: '', name: '', ruc: '', activity: '', address: '' }
export const parameters: UtilityParameters = { income: 0, legalPercent: 10, averageEmployees: 27, daysBase: 360, remunerationMonths: 12, capMonths: 18 }
export const companyClients: CompanyClient[] = [
  {
    id: '1',
    legalName: 'EMPRESA DEMO SAC',
    tradeName: 'EMPRESA DEMO SAC',
    ruc: '20123456789',
    activity: 'Comercio y restaurantes',
    contact: '—',
    role: 'Contacto principal',
    phone: '',
    email: '',
    address: 'Av. Los Héroes 123',
    status: 'Activo',
    employees: 0,
    lastExercise: 2026,
  },
]
export const employees: Employee[] = []
export const exercises: Exercise[] = []
export const auditLogs: AuditLog[] = []
export const versions: CalculationVersion[] = []
