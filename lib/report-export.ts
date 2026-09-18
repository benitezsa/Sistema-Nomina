'use client'

import type * as ExcelJS from 'exceljs'
import type { AuditLog, CalculationVersion, Company, EmployeeUtilityResult, UtilityParameters } from '@/components/utilities/types'
import { TRAMOS_QUINTA_CATEGORIA, UIT_DEDUCCION, UIT_DEFAULT, getTotals, money, number } from '@/lib/utilities-calculation'

export type ReportContext = {
  company: Company
  parameters: UtilityParameters
  results: EmployeeUtilityResult[]
  totals: ReturnType<typeof getTotals>
  year?: number
}

export type ReportHeader = {
  company: Company
  parameters: UtilityParameters
  year?: number
}

const MONEY_FMT = '"S/ " #,##0.00'
const BORDER_THIN = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
} as const
const HEADER_FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F3FF' } }

async function loadExcel() {
  return import('exceljs')
}

function today(): string {
  return new Date().toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' })
}

async function downloadWorkbook(wb: ExcelJS.Workbook, filename: string): Promise<void> {
  const buffer = await wb.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

function sectionTitle(ws: ExcelJS.Worksheet, row: number, text: string, cols: number, fill: boolean = false): number {
  ws.mergeCells(row, 1, row, cols)
  const c = ws.getCell(row, 1)
  c.value = text
  c.font = { bold: true, size: 12, color: { argb: 'FF6D28D9' } }
  if (fill) {
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F3FF' } }
    c.border = BORDER_THIN
  }
  return row + 1
}

function kv(ws: ExcelJS.Worksheet, row: number, label: string, value: number | string, kind: 'money' | 'number' | 'text', cols: number = 4): void {
  ws.getCell(row, 1).value = label
  ws.getCell(row, 1).font = { bold: true, color: { argb: 'FF374151' } }
  const v = ws.getCell(row, 2)
  if (kind === 'money') {
    v.value = value as number
    v.numFmt = MONEY_FMT
  } else if (kind === 'number') {
    v.value = value as number
    v.numFmt = '#,##0'
  } else {
    v.value = value as string
  }
  v.font = { color: { argb: 'FF111827' } }
  ws.mergeCells(row, 2, row, cols)
}

function writeTable(
  ws: ExcelJS.Worksheet,
  headers: string[],
  rows: (string | number)[][],
  moneyCols: number[],
  widths: number[]
): number {
  widths.forEach((w, i) => {
    ws.getColumn(i + 1).width = w
  })
  const start = ws.lastRow ? ws.lastRow.number + 1 : 1
  const headerRow = ws.getRow(start)
  headers.forEach((h, i) => {
    const c = headerRow.getCell(i + 1)
    c.value = h
    c.font = { bold: true, color: { argb: 'FF4C1D95' } }
    c.fill = HEADER_FILL
    c.border = BORDER_THIN
  })
  let row = start + 1
  for (const data of rows) {
    data.forEach((v, j) => {
      const c = ws.getCell(row, j + 1)
      c.border = BORDER_THIN
      if (moneyCols.includes(j)) {
        c.value = v as number
        c.numFmt = MONEY_FMT
      } else {
        c.value = v
      }
    })
    row++
  }
  return row
}

function fundBlock(ws: ExcelJS.Worksheet, ctx: ReportContext, cols: number, title = 'Reporte consolidado de utilidades'): number {
  const { parameters, totals, results } = ctx
  const year = ctx.year ?? new Date().getFullYear()
  let row = 1
  const { company } = ctx
  ws.mergeCells(row, 1, row, cols)
  const t = ws.getCell(row, 1)
  t.value = title
  t.font = { bold: true, size: 16, color: { argb: 'FF6D28D9' } }
  ws.getRow(row).height = 26
  row++
  ws.mergeCells(row, 1, row, cols)
  ws.getCell(row, 1).value = company.name ? `${company.name} · RUC ${company.ruc}` : 'Empresa no configurada'
  ws.getCell(row, 1).font = { bold: true, size: 11 }
  row++
  ws.mergeCells(row, 1, row, cols)
  ws.getCell(row, 1).value = `Ejercicio ${year} · Fondo de utilidades: ${money.format(totals.fund)} · ${results.length} trabajadores`
  ws.getCell(row, 1).font = { italic: true, color: { argb: 'FF6B7280' } }
  row++
  row++
  row = sectionTitle(ws, row, 'Fondo y parámetros', cols, true)
  kv(ws, row, 'Renta neta anual', parameters.income, 'money', cols); row++
  kv(ws, row, 'Porcentaje legal', `${parameters.legalPercent} %`, 'text', cols); row++
  kv(ws, row, 'Fondo de utilidades (50% días / 50% remuneraciones)', totals.fund, 'money', cols); row++
  kv(ws, row, 'Promedio de trabajadores', parameters.averageEmployees, 'number', cols); row++
  kv(ws, row, 'Días base', parameters.daysBase, 'number', cols); row++
  kv(ws, row, 'Meses de remuneración', parameters.remunerationMonths, 'number', cols); row++
  kv(ws, row, 'Tope máximo', `${parameters.capMonths} remuneraciones`, 'text', cols); row++
  row++
  row = sectionTitle(ws, row, 'Resultados de la distribución', cols, true)
  kv(ws, row, 'Total distribuido', totals.distributed, 'money', cols); row++
  kv(ws, row, 'Quinta categoría total', totals.totalQuintaCategoria, 'money', cols); row++
  kv(ws, row, 'Remanente (sobre tope)', totals.remainder, 'money', cols); row++
  kv(ws, row, 'Trabajadores con tope aplicado', totals.capped, 'number', cols); row++
  kv(ws, row, 'Trabajadores considerados', results.length, 'number', cols); row++
  row++
  return row
}

async function buildConsolidado(ExcelJS: typeof import('exceljs'), ctx: ReportContext): Promise<ExcelJS.Workbook> {
  const { Workbook } = ExcelJS
  const wb = new Workbook()
  const ws = wb.addWorksheet('Consolidado')
  let row = fundBlock(ws, ctx, 4, 'Reporte consolidado de utilidades')
  row = sectionTitle(ws, row, 'Detalle por trabajador', 7, true)
  const headers = ['Trabajador', 'Días efectivos', 'Rem. computable', 'Utilidad bruta', 'Quinta categoría', 'Utilidad neta', 'Excedente']
  const rows = ctx.results.map((r) => [
    `${r.name} (${r.code})`,
    r.diasEfectivos,
    r.remuneracionComputable,
    r.distribucion.utilidadBruta,
    r.distribucion.quintaCategoria,
    r.distribucion.utilidadNeta,
    r.remainder,
  ])
  const next = writeTable(ws, headers, rows, [2, 3, 4, 5, 6], [34, 14, 16, 16, 16, 16, 16])
  ws.mergeCells(next, 1, next, 6)
  ws.getCell(next, 1).value = `Total distribuido: ${money.format(ctx.totals.distributed)}`
  ws.getCell(next, 1).font = { bold: true }
  return wb
}

async function buildIndividual(ExcelJS: typeof import('exceljs'), ctx: ReportContext): Promise<ExcelJS.Workbook> {
  const { Workbook } = ExcelJS
  const wb = new Workbook()
  const ws = wb.addWorksheet('Detalle por trabajador')
  let row = fundBlock(ws, ctx, 12, 'Reporte individual de utilidades')
  row = sectionTitle(ws, row, 'Matriz de cálculo por trabajador', 12, true)
  const headers = [
    'Trabajador',
    'Días efectivos',
    'Factor días (%)',
    'Rem. computable',
    'Factor rem. (%)',
    'Util. por días',
    'Util. por rem.',
    'Utilidad bruta',
    'Tope',
    'Quinta categoría',
    'Excedente',
    'Utilidad neta',
  ]
  const rows = ctx.results.map((r) => [
    `${r.name} (${r.code})`,
    r.diasEfectivos,
    Math.round(r.daysFactor * 10000) / 100,
    r.remuneracionComputable,
    Math.round(r.remunerationFactor * 10000) / 100,
    r.distribucion.utilidadPorDias,
    r.distribucion.utilidadPorRemuneraciones,
    r.distribucion.utilidadBruta,
    r.cap,
    r.distribucion.quintaCategoria,
    r.remainder,
    r.distribucion.utilidadNeta,
  ])
  const next = writeTable(ws, headers, rows, [3, 5, 6, 7, 8, 9, 10, 11], [30, 12, 13, 14, 12, 15, 15, 15, 15, 15, 15, 15])
  ws.mergeCells(next, 1, next, 11)
  ws.getCell(next, 1).value = `Total distribuido: ${money.format(ctx.totals.distributed)}`
  ws.getCell(next, 1).font = { bold: true }
  return wb
}

async function buildMatriz(ExcelJS: typeof import('exceljs'), ctx: ReportContext): Promise<ExcelJS.Workbook> {
  const { Workbook } = ExcelJS
  const wb = new Workbook()
  const ws = wb.addWorksheet('Matriz de cálculo')
  let row = fundBlock(ws, ctx, 6, 'Matriz de cálculo de utilidades')
  row = sectionTitle(ws, row, 'Tramos de quinta categoría (UIT ' + number.format(UIT_DEFAULT) + ')', 6, true)
  kv(ws, row, 'Deducción legal', `${UIT_DEDUCCION} UIT`, 'text', 6); row++
  const tramoTable = TRAMOS_QUINTA_CATEGORIA.map((tr) => [
    money.format(tr.desde * UIT_DEFAULT),
    tr.hasta === null ? 'En adelante' : money.format(tr.hasta * UIT_DEFAULT),
    `${(tr.tasa * 100).toFixed(0)} %`,
  ])
  row = writeTable(ws, ['Tramo desde (S/)', 'Tramo hasta (S/)', 'Tasa'], tramoTable, [], [24, 24, 12])
  row++
  row = sectionTitle(ws, row, 'Fórmulas aplicadas', 6, true)
  kv(ws, row, 'Fondo', 'Renta neta × % legal', 'text', 6); row++
  kv(ws, row, 'Mitad por días', '50% × (días efectivos / total días)', 'text', 6); row++
  kv(ws, row, 'Mitad por remuneraciones', '50% × (rem. computable / total rem.)', 'text', 6); row++
  kv(ws, row, 'Preliminar', 'Mitad por días + mitad por remuneraciones', 'text', 6); row++
  kv(ws, row, 'Tope', `Remuneración × ${ctx.parameters.capMonths} meses`, 'text', 6); row++
  kv(ws, row, 'Final', 'MIN(preliminar; tope)', 'text', 6); row++
  kv(ws, row, 'Quinta categoría', 'Final × tasa marginal', 'text', 6); row++
  kv(ws, row, 'Utilidad neta', 'Final − quinta categoría', 'text', 6); row++
  row++
  row = sectionTitle(ws, row, 'Matriz por trabajador', 12, true)
  const headers = [
    'Trabajador',
    'Días efectivos',
    'Factor días (%)',
    'Rem. computable',
    'Factor rem. (%)',
    'Preliminar',
    'Tope',
    'Final',
    'Excedente',
    'Quinta categoría',
    'Utilidad neta',
  ]
  const rows = ctx.results.map((r) => [
    `${r.name} (${r.code})`,
    r.diasEfectivos,
    Math.round(r.daysFactor * 10000) / 100,
    r.remuneracionComputable,
    Math.round(r.remunerationFactor * 10000) / 100,
    r.preliminary,
    r.cap,
    r.finalAmount,
    r.remainder,
    r.distribucion.quintaCategoria,
    r.distribucion.utilidadNeta,
  ])
  const next = writeTable(ws, headers, rows, [3, 5, 6, 7, 8, 9, 10], [30, 12, 13, 14, 12, 14, 14, 14, 14, 14, 14])
  ws.mergeCells(next, 1, next, 10)
  ws.getCell(next, 1).value = `Total distribuido: ${money.format(ctx.totals.distributed)}`
  ws.getCell(next, 1).font = { bold: true }
  return wb
}

async function buildResultados(ExcelJS: typeof import('exceljs'), ctx: ReportContext): Promise<ExcelJS.Workbook> {
  const { Workbook } = ExcelJS
  const wb = new Workbook()
  const ws = wb.addWorksheet('Resultados')
  let row = fundBlock(ws, ctx, 9, 'Reporte de resultados de distribución')
  const sorted = [...ctx.results].sort((a, b) => b.distribucion.utilidadNeta - a.distribucion.utilidadNeta)
  row = sectionTitle(ws, row, 'Ranking de resultados', 9, true)
  const headers = ['#', 'Trabajador', 'Departamento', 'Días efectivos', 'Rem. computable', 'Utilidad bruta', 'Quinta categoría', 'Utilidad neta', 'Tope']
  const rows = sorted.map((r, i) => [
    i + 1,
    r.name,
    r.department,
    r.diasEfectivos,
    r.remuneracionComputable,
    r.distribucion.utilidadBruta,
    r.distribucion.quintaCategoria,
    r.distribucion.utilidadNeta,
    r.capApplied ? 'Aplicado' : '—',
  ])
  const next = writeTable(ws, headers, rows, [4, 5, 6, 7], [6, 26, 16, 12, 14, 15, 15, 15, 12])
  ws.mergeCells(next, 1, next, 8)
  ws.getCell(next, 1).value = `Total distribuido: ${money.format(ctx.totals.distributed)} · Remanente: ${money.format(ctx.totals.remainder)}`
  ws.getCell(next, 1).font = { bold: true }
  return wb
}

async function buildBoleta(ExcelJS: typeof import('exceljs'), header: ReportHeader, result: EmployeeUtilityResult): Promise<ExcelJS.Workbook> {
  const { Workbook } = ExcelJS
  const wb = new Workbook()
  const ws = wb.addWorksheet('Boleta')
  const { company, parameters } = header
  const year = header.year ?? new Date().getFullYear()
  let row = 1
  ws.mergeCells(row, 1, row, 4)
  ws.getCell(row, 1).value = 'Constancia de utilidades'
  ws.getCell(row, 1).font = { bold: true, size: 16, color: { argb: 'FF6D28D9' } }
  row++
  ws.mergeCells(row, 1, row, 4)
  ws.getCell(row, 1).value = company.name ? `${company.name} · RUC ${company.ruc}` : 'Empresa no configurada'
  row++
  ws.mergeCells(row, 1, row, 4)
  ws.getCell(row, 1).value = `Ejercicio ${year} · Generado el ${today()}`
  ws.getCell(row, 1).font = { italic: true, color: { argb: 'FF6B7280' } }
  row++
  row++
  row = sectionTitle(ws, row, 'Datos del trabajador', 4, true)
  kv(ws, row, 'Trabajador', `${result.name}`, 'text', 4); row++
  kv(ws, row, 'Código', result.code, 'text', 4); row++
  kv(ws, row, 'Cargo', result.role, 'text', 4); row++
  kv(ws, row, 'Departamento', result.department, 'text', 4); row++
  row++
  row = sectionTitle(ws, row, 'Datos base del cálculo', 4, true)
  kv(ws, row, 'Días laborables', result.diasLaborables ?? 0, 'number', 4); row++
  kv(ws, row, 'Días no laborados', result.diasNoLaborados ?? 0, 'number', 4); row++
  kv(ws, row, 'Días efectivos', result.diasEfectivos, 'number', 4); row++
  kv(ws, row, 'Remuneración computable', result.remuneracionComputable, 'money', 4); row++
  row++
  row = sectionTitle(ws, row, 'Distribución (50% / 50%)', 4, true)
  kv(ws, row, 'Utilidad por días efectivos', result.distribucion.utilidadPorDias, 'money', 4); row++
  kv(ws, row, 'Utilidad por remuneraciones', result.distribucion.utilidadPorRemuneraciones, 'money', 4); row++
  kv(ws, row, 'Utilidad bruta', result.distribucion.utilidadBruta, 'money', 4); row++
  if (result.capApplied) {
    kv(ws, row, `Tope aplicado (${parameters.capMonths} remuneraciones)`, result.cap, 'money', 4); row++
    kv(ws, row, 'Excedente', result.remainder, 'money', 4); row++
  }
  row++
  row = sectionTitle(ws, row, 'Quinta categoría', 4, true)
  kv(ws, row, 'Remuneración anual', result.quintaCategoria?.remuneracionAnual ?? result.distribucion.utilidadBruta, 'money', 4); row++
  kv(ws, row, 'Deducción 7 UIT', result.quintaCategoria?.deduccion7UIT ?? 0, 'money', 4); row++
  kv(ws, row, 'Base imponible', result.quintaCategoria?.baseImponible ?? 0, 'money', 4); row++
  kv(ws, row, 'Impuesto quinta categoría', result.distribucion.quintaCategoria, 'money', 4); row++
  row++
  ws.mergeCells(row, 1, row, 3)
  ws.getCell(row, 1).value = 'Total a pagar al trabajador'
  ws.getCell(row, 1).font = { bold: true, size: 12 }
  const total = ws.getCell(row, 4)
  total.value = result.distribucion.utilidadNeta
  total.numFmt = MONEY_FMT
  total.font = { bold: true, size: 12, color: { argb: 'FF6D28D9' } }
  return wb
}

async function buildBitacora(ExcelJS: typeof import('exceljs'), logs: AuditLog[], versions: CalculationVersion[]): Promise<ExcelJS.Workbook> {
  const { Workbook } = ExcelJS
  const wb = new Workbook()
  const ws = wb.addWorksheet('Bitácora')
  ws.mergeCells(1, 1, 1, 6)
  ws.getCell(1, 1).value = 'Bitácora de auditoría'
  ws.getCell(1, 1).font = { bold: true, size: 16, color: { argb: 'FF6D28D9' } }
  ws.mergeCells(2, 1, 2, 6)
  ws.getCell(2, 1).value = `Generado el ${today()}`
  ws.getCell(2, 1).font = { italic: true, color: { argb: 'FF6B7280' } }
  let row = 4
  row = sectionTitle(ws, row, 'Actividad reciente', 6, true)
  const logsRows = logs.map((l) => [l.id, l.date, l.user, l.action, l.detail, l.status])
  const next = writeTable(ws, ['ID', 'Fecha', 'Usuario', 'Acción', 'Detalle', 'Estado'], logsRows, [], [12, 18, 16, 20, 44, 12])
  row = next + 1
  row = sectionTitle(ws, row, 'Versiones del cálculo', 6, true)
  const versionsRows = versions.map((v) => [v.id, v.version, v.date, v.user, v.reason, v.total])
  writeTable(ws, ['ID', 'Versión', 'Fecha', 'Usuario', 'Motivo', 'Total'], versionsRows, [5], [12, 12, 18, 16, 44, 16])
  return wb
}

export async function exportConsolidado(ctx: ReportContext): Promise<void> {
  const wb = await buildConsolidado(await loadExcel(), ctx)
  await downloadWorkbook(wb, `reporte-consolidado-${ctx.year ?? new Date().getFullYear()}.xlsx`)
}

export async function exportIndividual(ctx: ReportContext): Promise<void> {
  const wb = await buildIndividual(await loadExcel(), ctx)
  await downloadWorkbook(wb, `reporte-individual-${ctx.year ?? new Date().getFullYear()}.xlsx`)
}

export async function exportMatriz(ctx: ReportContext): Promise<void> {
  const wb = await buildMatriz(await loadExcel(), ctx)
  await downloadWorkbook(wb, `matriz-calculo-${ctx.year ?? new Date().getFullYear()}.xlsx`)
}

export async function exportResultados(ctx: ReportContext): Promise<void> {
  const wb = await buildResultados(await loadExcel(), ctx)
  await downloadWorkbook(wb, `resultados-distribucion-${ctx.year ?? new Date().getFullYear()}.xlsx`)
}

export async function exportBoleta(header: ReportHeader, result: EmployeeUtilityResult): Promise<void> {
  const wb = await buildBoleta(await loadExcel(), header, result)
  await downloadWorkbook(wb, `constancia-${result.code}-${header.year ?? new Date().getFullYear()}.xlsx`)
}

export async function exportBoletas(ctx: ReportContext): Promise<void> {
  const { Workbook } = await loadExcel()
  const wb = new Workbook()
  for (const result of ctx.results) {
    const sheet = wb.addWorksheet((result.code || result.name).slice(0, 31))
    const built = await buildBoleta(await loadExcel(), { company: ctx.company, parameters: ctx.parameters, year: ctx.year }, result)
    const source = built.getWorksheet('Boleta')
    if (source) {
      source.eachRow({ includeEmpty: false }, (r, n) => {
        const destRow = sheet.getRow(n)
        r.eachCell({ includeEmpty: true }, (c, col) => {
          const dest = destRow.getCell(col)
          dest.value = c.value
          dest.font = c.font
          dest.fill = c.fill
          dest.border = c.border
          dest.numFmt = c.numFmt
          dest.alignment = c.alignment
        })
        destRow.height = r.height
      })
      const widthCount = Math.max(sheet.columnCount, 4)
      for (let i = 1; i <= widthCount; i++) {
        const w = source.getColumn(i).width
        if (w) sheet.getColumn(i).width = w
      }
    }
  }
  await downloadWorkbook(wb, `boletas-individuales-${ctx.year ?? new Date().getFullYear()}.xlsx`)
}

export async function exportBitacora(logs: AuditLog[], versions: CalculationVersion[]): Promise<void> {
  const wb = await buildBitacora(await loadExcel(), logs, versions)
  await downloadWorkbook(wb, `bitacora-auditoria-${new Date().getFullYear()}.xlsx`)
}

const esc = (v: unknown): string =>
  String(v ?? '').replace(/[&<>"']/g, (m) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[m]!
  )

function htmlTable(headers: string[], rows: (string | number | null | undefined)[][]): string {
  return `<table><thead><tr>${headers.map((h) => `<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows
    .map(
      (r) =>
        `<tr>${r
          .map((c) => `<td${c === null || c === undefined ? ' class="muted"' : ''}>${c === null || c === undefined ? '—' : c}</td>`)
          .join('')}</tr>`
    )
    .join('')}</tbody></table>`
}

function htmlHeader(ctx: ReportContext, title: string): string {
  const { company } = ctx
  return `<div class="doc-header"><div><p class="eyebrow">Utilidades · Centro de reportes</p><h1>${esc(title)}</h1><p class="sub">Ejercicio ${ctx.year ?? new Date().getFullYear()} · Generado el ${today()}</p></div>${company.name ? `<div class="company"><p class="cname">${esc(company.name)}</p><p class="cmeta">RUC ${esc(company.ruc)}</p></div>` : ''}</div>`
}

function htmlKpis(ctx: ReportContext): string {
  const { totals, results } = ctx
  return `<section class="kpis"><div><p>Fondo de utilidades</p><strong>${esc(money.format(totals.fund))}</strong></div><div><p>Total distribuido</p><strong>${esc(money.format(totals.distributed))}</strong></div><div><p>Quinta categoría</p><strong>${esc(money.format(totals.totalQuintaCategoria))}</strong></div><div><p>Remanente</p><strong>${esc(money.format(totals.remainder))}</strong></div><div><p>Trabajadores</p><strong>${results.length}</strong></div></section>`
}

export function buildConsolidadoHtml(ctx: ReportContext): string {
  const sorted = [...ctx.results].sort((a, b) => b.distribucion.utilidadNeta - a.distribucion.utilidadNeta)
  const rows = sorted.map((r) => [
    `${esc(r.name)}<div class="sub">${esc(r.code)} · ${esc(r.department)}</div>`,
    number.format(r.diasEfectivos),
    money.format(r.remuneracionComputable),
    money.format(r.distribucion.utilidadBruta),
    money.format(r.distribucion.quintaCategoria),
    money.format(r.distribucion.utilidadNeta),
  ])
  return `${htmlHeader(ctx, 'Reporte consolidado')}${htmlKpis(ctx)}<section><h2>Detalle por trabajador</h2>${htmlTable(['Trabajador', 'Días efectivos', 'Rem. computable', 'Utilidad bruta', 'Quinta categoría', 'Utilidad neta'], rows)}</section>`
}

export function buildIndividualHtml(ctx: ReportContext): string {
  const sorted = [...ctx.results].sort((a, b) => b.distribucion.utilidadNeta - a.distribucion.utilidadNeta)
  const rows = sorted.map((r) => [
    `${esc(r.name)}<div class="sub">${esc(r.code)} · ${esc(r.department)}</div>`,
    number.format(r.diasEfectivos),
    `${(r.daysFactor * 100).toFixed(2)}%`,
    money.format(r.remuneracionComputable),
    `${(r.remunerationFactor * 100).toFixed(2)}%`,
    money.format(r.distribucion.utilidadPorDias),
    money.format(r.distribucion.utilidadPorRemuneraciones),
    money.format(r.distribucion.utilidadBruta),
    r.capApplied ? money.format(r.cap) : '—',
    money.format(r.distribucion.quintaCategoria),
    money.format(r.distribucion.utilidadNeta),
  ])
  return `${htmlHeader(ctx, 'Reporte individual')}${htmlKpis(ctx)}<section><h2>Detalle de cálculo por trabajador</h2>${htmlTable(['Trabajador', 'Días efectivos', 'Factor días', 'Rem. computable', 'Factor rem.', 'Util. días', 'Util. rem.', 'Util. bruta', 'Tope', 'Quinta cat.', 'Util. neta'], rows)}</section>`
}

export function buildMatrizHtml(ctx: ReportContext): string {
  const { parameters } = ctx
  const sorted = [...ctx.results].sort((a, b) => b.distribucion.utilidadNeta - a.distribucion.utilidadNeta)
  const tramos = TRAMOS_QUINTA_CATEGORIA.map((tr) => [
    money.format(tr.desde * UIT_DEFAULT),
    tr.hasta === null ? 'En adelante' : money.format(tr.hasta * UIT_DEFAULT),
    `${(tr.tasa * 100).toFixed(0)}%`,
  ])
  const rows = sorted.map((r) => [
    `${esc(r.name)}<div class="sub">${esc(r.code)} · ${esc(r.department)}</div>`,
    number.format(r.diasEfectivos),
    `${(r.daysFactor * 100).toFixed(2)}%`,
    money.format(r.remuneracionComputable),
    `${(r.remunerationFactor * 100).toFixed(2)}%`,
    money.format(r.preliminary),
    money.format(r.cap),
    money.format(r.finalAmount),
    r.remainder > 0 ? money.format(r.remainder) : '—',
    money.format(r.distribucion.quintaCategoria),
    money.format(r.distribucion.utilidadNeta),
  ])
  return `${htmlHeader(ctx, 'Matriz de cálculo')}${htmlKpis(ctx)}<section><h2>Parámetros y fórmulas</h2><table class="kv"><tbody><tr><td>Renta neta</td><td>${esc(money.format(parameters.income))}</td></tr><tr><td>% legal</td><td>${parameters.legalPercent} %</td></tr><tr><td>Tope máximo</td><td>${parameters.capMonths} remuneraciones</td></tr><tr><td>Fórmula</td><td>Final = MIN(preliminar; tope); Neta = Final − quinta categoría</td></tr></tbody></table></section><section><h2>Tramos de quinta categoría (UIT ${esc(number.format(UIT_DEFAULT))})</h2>${htmlTable(['Tramo desde (S/)', 'Tramo hasta (S/)', 'Tasa'], tramos)}</section><section><h2>Matriz por trabajador</h2>${htmlTable(['Trabajador', 'Días', 'Factor días', 'Rem. comp.', 'Factor rem.', 'Preliminar', 'Tope', 'Final', 'Excedente', 'Quinta cat.', 'Neta'], rows)}</section>`
}

export function buildBoletaHtml(header: ReportHeader, result: EmployeeUtilityResult): string {
  const { company, parameters } = header
  const fila = (label: string, valor: string, extra?: string) =>
    `<div class="fila"><span>${esc(label)}</span><span><strong>${valor}</strong>${extra ? `<span class="accent">${esc(extra)}</span>` : ''}</span></div>`
  return `<div class="doc-header"><div><p class="eyebrow">Constancia de utilidades</p><h1>Boleta individual de distribución</h1><p class="sub">Ejercicio ${header.year ?? new Date().getFullYear()}</p></div>${company.name ? `<div class="company"><p class="cname">${esc(company.name)}</p><p class="cmeta">RUC ${esc(company.ruc)}</p></div>` : ''}</div><section class="worker"><div><p class="lbl">Trabajador</p><p><strong>${esc(result.name)}</strong></p><p class="sub">${esc(result.code)} · ${esc(result.department)}</p></div><div><p class="lbl">Cargo</p><p><strong>${esc(result.role)}</strong></p></div><div><p class="lbl">Jornada</p><p><strong>${result.jornada ? `${result.jornada} días/semana` : '6 días/semana'}</strong></p></div></section><section><h2>Datos base del cálculo</h2>${fila('Días laborables', number.format(result.diasLaborables ?? 0))}${fila('Días no laborados', number.format(result.diasNoLaborados ?? 0))}${fila('Días efectivos', number.format(result.diasEfectivos), `factor ${(result.daysFactor * 100).toFixed(2)}%`)}${fila('Remuneración computable', money.format(result.remuneracionComputable), `factor ${(result.remunerationFactor * 100).toFixed(2)}%`)}</section><section><h2>Distribución (50% / 50%)</h2>${fila('Utilidad por días efectivos', money.format(result.distribucion.utilidadPorDias))}${fila('Utilidad por remuneraciones', money.format(result.distribucion.utilidadPorRemuneraciones))}${fila('Utilidad bruta', money.format(result.distribucion.utilidadBruta))}${result.capApplied ? fila(`Tope aplicado (${parameters.capMonths} remuneraciones)`, money.format(result.cap)) + fila('Excedente', money.format(result.remainder)) : ''}</section><section><h2>Quinta categoría</h2>${fila('Remuneración anual', money.format(result.quintaCategoria?.remuneracionAnual ?? result.distribucion.utilidadBruta))}${fila('Deducción 7 UIT', `- ${money.format(result.quintaCategoria?.deduccion7UIT ?? 0)}`)}${fila('Base imponible', money.format(result.quintaCategoria?.baseImponible ?? 0))}${fila('Impuesto quinta categoría', money.format(result.distribucion.quintaCategoria))}</section><div class="total"><div><p>Total a pagar al trabajador</p><p class="sub">Utilidad bruta menos quinta categoría</p></div><strong>${esc(money.format(result.distribucion.utilidadNeta))}</strong></div>`
}

export function openPrintWindow(title: string, body: string): void {
  const w = window.open('', '_blank', 'width=980,height=760')
  if (!w) return
  w.document.write(`<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><title>${esc(title)}</title><style>
html,body{margin:0;padding:0;font-family:system-ui,sans-serif;color:#111827}
.toolbar{position:sticky;top:0;display:flex;justify-content:flex-end;gap:8px;padding:10px 16px;background:#6D28D9;z-index:9}
.toolbar button{border:1px solid rgba(255,255,255,.6);background:transparent;color:#fff;font-weight:600;font-size:13px;padding:6px 14px;border-radius:8px;cursor:pointer}
.toolbar button:hover{background:rgba(255,255,255,.12)}
.content{max-width:960px;margin:0 auto;padding:28px 24px 60px}
.doc-header{display:flex;justify-content:space-between;gap:16px;border-bottom:2px solid #6D28D9;padding-bottom:14px;margin-bottom:18px}
.eyebrow{font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:#6D28D9;font-weight:700;margin:0 0 4px}
h1{margin:0;font-size:22px}
.sub{color:#6B7280;font-size:12px;margin-top:4px}
.company{text-align:right}.cname{font-weight:700;margin:0}.cmeta{color:#6B7280;font-size:12px;margin:0}
.kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:20px}
.kpis div{border:1px solid #E5E7EB;border-radius:10px;padding:12px;background:#FAFAFF}
.kpis p{margin:0 0 4px;font-size:11px;color:#6B7280}.kpis strong{font-size:15px}
section{margin-bottom:22px}h2{font-size:14px;color:#6D28D9;margin:0 0 10px}
table{width:100%;border-collapse:collapse;font-size:12px}
th,td{border:1px solid #E5E7EB;padding:6px 8px;text-align:left;vertical-align:top}
th{background:#F5F3FF;color:#4C1D95;font-weight:600;font-size:11px}
td.muted{color:#9CA3AF}
tr:nth-child(even) td{background:#FBFAFE}
.sub{color:#6B7280;font-size:11px}
table.kv td:first-child{width:200px;color:#6B7280;font-weight:600;background:#F9FAFB}
.fila{display:flex;justify-content:space-between;border-bottom:1px solid #E5E7EB;padding:7px 0;font-size:13px}
.fila span:first-child{color:#6B7280}.accent{color:#6D28D9;padding-left:8px;font-size:12px}
.worker{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;border:1px solid #E5E7EB;border-radius:12px;padding:14px;margin-bottom:20px}
.lbl{font-size:10px;text-transform:uppercase;letter-spacing:.08em;color:#6B7280;margin:0 0 2px}
.total{display:flex;justify-content:space-between;align-items:center;border:1px solid #E5E7EB;border-radius:12px;background:#F5F3FF;padding:16px}
.total strong{font-size:22px;color:#6D28D9}
@media print{.toolbar{display:none}.content{padding:0}}
</style></head><body><div class="toolbar"><button onclick="window.print()">Imprimir</button><button onclick="window.close()">Cerrar</button></div><div class="content">${body}</div></body></html>`)
  w.document.close()
  w.focus()
}
