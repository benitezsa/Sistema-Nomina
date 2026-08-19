import type { AuditLog, CalculationVersion, Company, CompanyClient, Employee, Exercise, UtilityParameters } from '@/components/utilities/types'

export const company: Company = { id: 'comp-01', name: 'Inversiones Andinas S.A.C.', ruc: '20548796321', activity: 'Servicios financieros', address: 'Av. República de Panamá 3531, San Isidro, Lima' }
export const parameters: UtilityParameters = { income: 1500000, legalPercent: 10, averageEmployees: 27, daysBase: 360, remunerationMonths: 12, capMonths: 18 }
export const companyClients: CompanyClient[] = [
  { id: 'cli-01', legalName: 'Inversiones Andinas S.A.C.', tradeName: 'Andinas Finanzas', ruc: '20548796321', activity: 'Servicios financieros', contact: 'Mariana Flores', role: 'Gerente de Administración', phone: '+51 987 654 321', email: 'mariana.flores@andinasfinanzas.pe', address: 'Av. República de Panamá 3531, San Isidro, Lima', status: 'Activo', employees: 27, lastExercise: 2026 },
  { id: 'cli-02', legalName: 'Comercial del Pacífico S.A.', tradeName: 'Pacífico Comercial', ruc: '20481237659', activity: 'Comercio mayorista', contact: 'Luis Gamarra', role: 'Jefe de Recursos Humanos', phone: '+51 955 321 876', email: 'luis.gamarra@pacificocomercial.pe', address: 'Jr. Monterrey 145, Santiago de Surco, Lima', status: 'Activo', employees: 84, lastExercise: 2026 },
  { id: 'cli-03', legalName: 'Tecnología Río Verde E.I.R.L.', tradeName: 'Río Verde Tech', ruc: '20607894512', activity: 'Tecnología y software', contact: 'Paula Rojas', role: 'Administradora', phone: '+51 966 248 740', email: 'paula.rojas@rioverde.tech', address: 'Calle Las Begonias 441, San Isidro, Lima', status: 'Pendiente', employees: 19, lastExercise: 2025 },
  { id: 'cli-04', legalName: 'Constructora Horizonte S.A.C.', tradeName: 'Horizonte Obras', ruc: '20563981247', activity: 'Construcción', contact: 'Álvaro Medina', role: 'Director Financiero', phone: '+51 944 610 228', email: 'amedina@horizonteobras.pe', address: 'Av. El Derby 254, Santiago de Surco, Lima', status: 'Activo', employees: 132, lastExercise: 2026 },
  { id: 'cli-05', legalName: 'Servicios del Sur S.R.L.', tradeName: 'Sur Servicios', ruc: '20415987306', activity: 'Servicios empresariales', contact: 'Rosa Valdivia', role: 'Contadora General', phone: '+51 978 430 115', email: 'rvaldivia@surservicios.pe', address: 'Av. Ejército 710, Miraflores, Lima', status: 'Inactivo', employees: 0, lastExercise: 2023 },
]
export const employees: Employee[] = [
  { id: 'e-01', code: 'EMP-001', name: 'Ana María Torres', role: 'Gerente de Operaciones', department: 'Operaciones', days: 360, remuneration: 16800, status: 'Completo' },
  { id: 'e-02', code: 'EMP-002', name: 'Carlos Mendoza', role: 'Jefe de Finanzas', department: 'Finanzas', days: 360, remuneration: 14200, status: 'Completo' },
  { id: 'e-03', code: 'EMP-003', name: 'Lucía Ramírez', role: 'Especialista Senior', department: 'Comercial', days: 320, remuneration: 9800, status: 'Completo' },
  { id: 'e-04', code: 'EMP-004', name: 'Jorge Castillo', role: 'Analista de Riesgos', department: 'Riesgos', days: 280, remuneration: 7200, status: 'Observado' },
  { id: 'e-05', code: 'EMP-005', name: 'Valeria Salazar', role: 'Analista Contable', department: 'Finanzas', days: 360, remuneration: 6800, status: 'Completo' },
  { id: 'e-06', code: 'EMP-006', name: 'Diego Paredes', role: 'Asesor Comercial', department: 'Comercial', days: 190, remuneration: 5600, status: 'Completo' },
  { id: 'e-07', code: 'EMP-007', name: 'Sofía Herrera', role: 'Coordinadora Legal', department: 'Legal', days: 360, remuneration: 11500, status: 'Completo' },
  { id: 'e-08', code: 'EMP-008', name: 'Mateo Vargas', role: 'Especialista TI', department: 'Tecnología', days: 360, remuneration: 12600, status: 'Pendiente' },
  { id: 'e-09', code: 'EMP-009', name: 'Camila Núñez', role: 'Analista de Datos', department: 'Tecnología', days: 240, remuneration: 7600, status: 'Completo' },
  { id: 'e-10', code: 'EMP-010', name: 'Renato León', role: 'Asistente Administrativo', department: 'Administración', days: 360, remuneration: 4200, status: 'Completo' },
  { id: 'e-11', code: 'EMP-011', name: 'Mariana Flores', role: 'Consultora', department: 'Operaciones', days: 340, remuneration: 8900, status: 'Completo' },
  { id: 'e-12', code: 'EMP-012', name: 'Alonso Díaz', role: 'Técnico de Campo', department: 'Operaciones', days: 160, remuneration: 3900, status: 'Completo' },
]
export const exercises: Exercise[] = [
  { id: 'ej-2026', year: 2026, company, status: 'Calculado', createdAt: '02 ene 2026', updatedAt: '18 mar 2026', employeeCount: 27, fund: 150000 },
  { id: 'ej-2025', year: 2025, company, status: 'Cerrado', createdAt: '05 ene 2025', updatedAt: '28 mar 2025', employeeCount: 25, fund: 132500 },
  { id: 'ej-2024', year: 2024, company, status: 'Cerrado', createdAt: '04 ene 2024', updatedAt: '22 mar 2024', employeeCount: 24, fund: 118400 },
]
export const auditLogs: AuditLog[] = [
  { id: 'a1', date: '18 mar 2026, 16:42', user: 'Mariana Flores', action: 'Cálculo ejecutado', detail: 'Versión 1.2 · 27 trabajadores procesados', status: 'Completado' },
  { id: 'a2', date: '18 mar 2026, 16:18', user: 'Mariana Flores', action: 'Parámetros actualizados', detail: 'Renta neta modificada a S/ 1,500,000.00', status: 'Completado' },
  { id: 'a3', date: '17 mar 2026, 11:05', user: 'Luis Gamarra', action: 'Trabajadores importados', detail: 'Carga nómina-marzo-2026.xlsx · 27 registros', status: 'Completado' },
  { id: 'a4', date: '02 ene 2026, 09:12', user: 'Luis Gamarra', action: 'Ejercicio creado', detail: 'Ejercicio 2026 · Inversiones Andinas S.A.C.', status: 'Completado' },
]
export const versions: CalculationVersion[] = [
  { id: 'v12', version: '1.2', date: '18 mar 2026, 16:42', user: 'Mariana Flores', reason: 'Cálculo vigente', total: 150000 },
  { id: 'v11', version: '1.1', date: '18 mar 2026, 16:20', user: 'Mariana Flores', reason: 'Actualización de renta neta', total: 150000 },
  { id: 'v10', version: '1.0', date: '17 mar 2026, 11:35', user: 'Luis Gamarra', reason: 'Carga inicial de información', total: 147800 },
]
