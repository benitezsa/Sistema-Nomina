import ExcelJS from 'exceljs'

export const months = [
  'enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
  'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre',
] as const

export type Mes = (typeof months)[number]

export const MESES_ENCABEZADO: string[] = [
  'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
  'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre',
]

export const MESES_EN_CERO = {
  enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0,
  julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0,
  total: 0,
}

export function mensual(valores: Partial<Record<Mes, number>>) {
  const out = { ...MESES_EN_CERO }
  for (const [mes, valor] of Object.entries(valores)) out[mes as Mes] = valor
  out.total = months.reduce((suma, mes) => suma + out[mes], 0)
  return out
}

export function libro(hojas: { nombre: string; encabezados: unknown[]; filas: unknown[][] }[]) {
  const wb = new ExcelJS.Workbook()
  for (const hoja of hojas) {
    const ws = wb.addWorksheet(hoja.nombre)
    ws.addRow(hoja.encabezados)
    for (const fila of hoja.filas) ws.addRow(fila)
  }
  return wb
}

/** Libro de días: matriz mensual de días laborados por trabajador. */
export function libroDias(
  filas: { dni: string; nombres: string; dias: Partial<Record<Mes, number>> }[],
  { total = 'Total días trabajados' } = {},
) {
  return libro([{
    nombre: 'DIAS',
    encabezados: ['DNI', 'Nombres', ...MESES_ENCABEZADO, total],
    filas: filas.map((f) => [f.dni, f.nombres, ...months.map((mes) => f.dias[mes] ?? 0)]),
  }])
}

/** Libro de faltas: matriz mensual de días no laborados por trabajador. */
export function libroFaltas(
  filas: { dni: string; nombres: string; faltas: Partial<Record<Mes, number>> }[],
) {
  return libro([{
    nombre: 'FALTAS',
    encabezados: ['DNI', 'Nombres', ...MESES_ENCABEZADO, 'Total faltas'],
    filas: filas.map((f) => [f.dni, f.nombres, ...months.map((mes) => f.faltas[mes] ?? 0)]),
  }])
}

/**
 * Libro de días completo: la matriz de días laborados y la de días no
 * laborados. El archivo aporta días e incidencias en fuentes separadas.
 */
export function libroDiasYFaltas(
  filas: {
    dni: string
    nombres: string
    dias: Partial<Record<Mes, number>>
    faltas?: Partial<Record<Mes, number>>
  }[],
) {
  return libro([
    {
      nombre: 'DIAS',
      encabezados: ['DNI', 'Nombres', ...MESES_ENCABEZADO, 'Total días trabajados'],
      filas: filas.map((f) => [f.dni, f.nombres, ...months.map((mes) => f.dias[mes] ?? 0)]),
    },
    {
      nombre: 'FALTAS',
      encabezados: ['DNI', 'Nombres', ...MESES_ENCABEZADO, 'Total faltas'],
      filas: filas.map((f) => [f.dni, f.nombres, ...months.map((mes) => f.faltas?.[mes] ?? 0)]),
    },
  ])
}

/** Libro de remuneraciones: matriz mensual de importes por trabajador. */
export function libroRemuneraciones(
  filas: { dni: string; nombres: string; ingreso?: string; remuneraciones: Partial<Record<Mes, number>> }[],
) {
  return libro([{
    nombre: 'REM',
    encabezados: ['DNI', 'Nombres', 'Fecha ingreso', ...MESES_ENCABEZADO],
    filas: filas.map((f) => [
      f.dni,
      f.nombres,
      f.ingreso ?? '',
      ...months.map((mes) => f.remuneraciones[mes] ?? 0),
    ]),
  }])
}
