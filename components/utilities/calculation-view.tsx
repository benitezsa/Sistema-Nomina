  'use client'

  import { Fragment, useMemo, useRef, useState, useEffect } from 'react'
  import ExcelJS from 'exceljs'

  import { AlertCircle, ArrowDownToLine, ArrowRight, CalendarDays, Calculator, Check, ChevronRight, ClipboardCheck, Download, FileSpreadsheet, Filter, Play, Plus, Search, ShieldCheck, Trash2, Upload, Users, X } from 'lucide-react'
  import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert'
  import { Avatar, AvatarFallback } from '@/components/ui/avatar'
  import { Badge } from '@/components/ui/badge'
  import { Button } from '@/components/ui/button'
  import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
  import { Input } from '@/components/ui/input'
  import { Progress } from '@/components/ui/progress'
  import { Separator } from '@/components/ui/separator'
  import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from '@/components/ui/sheet'
  import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
  import { calculateFund, calcularDiasNoLaborados, getTotals, money, number } from '@/lib/utilities-calculation'
  import {
    consolidarTrabajadores,
    extraerDias,
    extraerRemuneraciones,
    formatearPeriodo,
    inferirAnioEjercicio,
    leerPeriodo,
    mesesTrabajados,
    type ImportedWorker as ImportedWorkerMotor,
    type IncidenciaImportada,
    type PeriodoIncidencia,
  } from '@/lib/import-utilidades'
  import type { ResumenRemuneracion } from '@/lib/import-utilidades-atomico'
  import { useSettings } from '@/lib/settings-context'
  import type { Employee, EmployeeUtilityResult, ImportSummary, Jornada } from './types'
  

  function SectionTitle({ eyebrow, title, description, action }: { eyebrow?: string; title: string; description?: string; action?: React.ReactNode }) { return <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between"><div><p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">{eyebrow}</p><h1 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">{title}</h1>{description && <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">{description}</p>}</div>{action}</div> }
  function Kpi({ label, value, hint, tone = 'default', icon: Icon }: { label: string; value: string; hint: string; tone?: 'default' | 'teal' | 'amber' | 'rose'; icon: typeof Calculator }) { return <Card className="border-border/70 shadow-sm"><CardContent className="flex items-start justify-between p-5"><div><p className="text-xs font-medium text-muted-foreground">{label}</p><p className="mt-2 text-xl font-semibold tracking-tight">{value}</p><p className={`mt-1 text-xs ${tone === 'rose' ? 'text-rose-400' : tone === 'amber' ? 'text-amber-400' : tone === 'teal' ? 'text-teal-400' : 'text-muted-foreground'}`}>{hint}</p></div><div className={`rounded-lg p-2.5 ${tone === 'teal' ? 'bg-teal-500/10 text-teal-400' : tone === 'amber' ? 'bg-amber-500/10 text-amber-400' : tone === 'rose' ? 'bg-rose-500/10 text-rose-400' : 'bg-primary/8 text-primary'}`}><Icon className="size-4" /></div>
  </CardContent></Card> }

  type MonthlyValues = {
    enero: number
    febrero: number
    marzo: number
    abril: number
    mayo: number
    junio: number
    julio: number
    agosto: number
    septiembre: number
    octubre: number
    noviembre: number
    diciembre: number
    total: number
  }

  type FeriadoImportado = {
  fecha: ExcelJS.CellValue
  descripcion: string
}

  // El trabajador normalizado es el que entrega el motor: la vista solo acota
  // la jornada a los valores válidos del dominio (4, 5 o 6 días por semana).
  type ImportedWorker = Omit<ImportedWorkerMotor, 'jornada'> & { jornada?: Jornada }


  function adaptarTrabajadoresMotor(trabajadores: ImportedWorkerMotor[]): ImportedWorker[] {
    return trabajadores.map((trabajador) => ({
      ...trabajador,
      jornada: trabajador.jornada === 4 || trabajador.jornada === 5 || trabajador.jornada === 6 ? trabajador.jornada : undefined,
    }))
  }

  const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']   as const

  const MONTH_LABELS: Record<string, string> = {
    enero: 'Enero', febrero: 'Febrero', marzo: 'Marzo', abril: 'Abril',
    mayo: 'Mayo', junio: 'Junio', julio: 'Julio', agosto: 'Agosto',
    septiembre: 'Septiembre', octubre: 'Octubre', noviembre: 'Noviembre', diciembre: 'Diciembre',
  }

  type EstadoPeriodo = 'Trabajó' | 'No trabajó' | 'Reingresó'

  const ESTADO_PERIODO_STYLE: Record<EstadoPeriodo, string> = {
    'Trabajó': 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30',
    'No trabajó': 'bg-slate-500/10 text-slate-400 border-slate-500/30',
    'Reingresó': 'bg-amber-500/10 text-amber-400 border-amber-500/30',
  }

  const formatDate = (date: Date) => {
    const day = String(date.getUTCDate()).padStart(2, '0')
    const month = String(date.getUTCMonth() + 1).padStart(2, '0')
    const year = date.getUTCFullYear()

    return `${day}/${month}/${year}`
  }

  const formatExcelDate = (value: ExcelJS.CellValue) => {
    if (!value) return '—'

    if (value instanceof Date) {
      return formatDate(value)
    }

    if (typeof value === 'object') {
      if ('result' in value) {
        return formatExcelDate(value.result as ExcelJS.CellValue)
      }

      if ('text' in value && typeof value.text === 'string') {
        return formatExcelDate(value.text)
      }
    }

    if (typeof value === 'number') {
      const excelEpoch = Date.UTC(1899, 11, 30)
      return formatDate(new Date(excelEpoch + value * 86_400_000))
    }

    if (typeof value === 'string') {
      const date = new Date(value)
      return Number.isNaN(date.getTime()) ? value : formatDate(date)
    }

    return '—'
  }

  // Los meses trabajados del trabajador salen siempre de mesesTrabajados (motor):
  // un mes cuenta como trabajado si el período tiene días registrados > 0 y tener
  // una fila en el archivo no lo convierte en mes trabajado. El segundo argumento
  // es la corrección manual de la permanencia en Validación.
  function getPeriodosTrabajados(worker: ImportedWorker, mesesOverride?: string[]): { mes: string; nombre: string; estado: EstadoPeriodo }[] {
    const mesesConDias = new Set(mesesTrabajados(worker, mesesOverride))
    const resultado: { mes: string; nombre: string; estado: EstadoPeriodo }[] = []
    let prevPresente = false
    let trabajóAntes = false

    for (const mes of months) {
      const presente = mesesConDias.has(mes)
      let estado: EstadoPeriodo

      if (presente) {
        estado = trabajóAntes && !prevPresente ? 'Reingresó' : 'Trabajó'
        trabajóAntes = true
        prevPresente = true
      } else {
        estado = 'No trabajó'
        prevPresente = false
      }

      resultado.push({ mes, nombre: MONTH_LABELS[mes], estado })
    }

    return resultado
  }

  const pad2 = (n: number) => String(n).padStart(2, '0')

  const formatInputDate = (value: ExcelJS.CellValue): string => {
    if (!value) return ''

    let date: Date | null = null

    if (value instanceof Date) {
      date = value
    } else if (typeof value === 'number') {
      date = new Date(Date.UTC(1899, 11, 30) + value * 86_400_000)
    } else if (typeof value === 'string') {
      const parsed = new Date(value)
      if (!Number.isNaN(parsed.getTime())) date = parsed
    } else if (typeof value === 'object') {
      if ('result' in value) return formatInputDate(value.result as ExcelJS.CellValue)
      if ('text' in value && typeof value.text === 'string') return formatInputDate(value.text)
    }

    if (!date || Number.isNaN(date.getTime())) return ''

    return `${date.getUTCFullYear()}-${pad2(date.getUTCMonth() + 1)}-${pad2(date.getUTCDate())}`
  }

  const parseIncFecha = (value: ExcelJS.CellValue): Date | null => {
    const s = formatInputDate(value)
    if (!s) return null
    const date = new Date(`${s}T00:00:00Z`)
    return Number.isNaN(date.getTime()) ? null : date
  }

  const SIN_DATO = 'No disponible'

  /**
   * Resuelve una fecha del trabajador a partir de todas las fuentes que la
   * aportan (edición de Validación primero, después el dato importado) y devuelve
   * una única versión normalizada. No inventa fechas: si ninguna fuente tiene
   * el dato, se informa "No disponible".
   */
  function resolverFecha(...candidatos: (ExcelJS.CellValue | string | undefined)[]): { iso: string; etiqueta: string } {
    for (const candidato of candidatos) {
      const iso = formatInputDate(candidato)
      if (!iso) continue
      const [anio, mes, dia] = iso.split('-')
      return { iso, etiqueta: `${dia}/${mes}/${anio}` }
    }
    return { iso: '', etiqueta: SIN_DATO }
  }

  /** Fecha de ingreso del trabajador: fuente única para toda la vista. */
  function fechaIngresoTrabajador(trabajador: ImportedWorker, editada?: string) {
    return resolverFecha(editada, trabajador.fechaInicio)
  }

  /** Fecha de cese del trabajador: misma fuente única que el ingreso. */
  function fechaCeseTrabajador(trabajador: ImportedWorker, editada?: string) {
    return resolverFecha(editada, trabajador.fechaCese)
  }

  function observacionesRemuneracionImportada(resumen?: ResumenRemuneracion): string[] {
    if (!resumen) return []

    const observaciones: string[] = []

    if (!resumen.mensualDisponible && resumen.totalDeclarado !== undefined) {
      observaciones.push('El archivo solo aporta total anual de remuneración: se usa ese importe y no se distribuyen meses.')
    }
    for (const conflicto of resumen.conflictos) observaciones.push(`Revisión de importes: ${conflicto}`)
    if (resumen.duplicados > 0) {
      observaciones.push(`Se descartaron ${resumen.duplicados} registro(s) duplicado(s) para no duplicar la remuneración.`)
    }
    if (resumen.conceptosDesconocidos.length > 0) {
      observaciones.push(`Conceptos fuera del catálogo, excluidos del cálculo: ${resumen.conceptosDesconocidos.join(', ')}.`)
    }
    if (resumen.conceptosExcluidos.length > 0) {
      observaciones.push(`Conceptos no remunerativos, excluidos del cálculo: ${resumen.conceptosExcluidos.join(', ')}.`)
    }
    if (resumen.lotesAnteriores.length > 0) {
      observaciones.push(`Este lote reemplaza a: ${resumen.lotesAnteriores.join(', ')} (se conserva la trazabilidad).`)
    }

    return observaciones
  }

  function getObservaciones(trabajador: ImportedWorker): string[] {
    const observaciones: string[] = []

    if (trabajador.remuneraciones.total <= 0) observaciones.push('Sin remuneraciones registradas para el ejercicio.')
    if (trabajador.remuneracionIncompleta) observaciones.push('Remuneración anual pendiente: el archivo no contiene los 12 meses necesarios para calcularla.')
    if ((trabajador.diasTrabajados?.total ?? 0) <= 0) observaciones.push('Sin días laborados registrados para el ejercicio.')
    if (!fechaIngresoTrabajador(trabajador).iso) observaciones.push('Fecha de ingreso pendiente de registrar.')
    if (!`${trabajador.apellidoPaterno}${trabajador.apellidoMaterno}${trabajador.nombres}`.trim()) observaciones.push('Nombres incompletos en la base importada.')
    const pendientes = pendientesDeRevision(trabajador.incidencias ?? [])
    if (pendientes.length > 0) {
      observaciones.push(`${pendientes.length} incidencia(s) sin identificar tipo o importe: quedan pendientes de revisión y no descuentan días.`)
    }
    observaciones.push(...observacionesRemuneracionImportada(trabajador.remuneracionResumen))

    return observaciones
  }

  type EstadoValidacion = 'Completo' | 'Pendiente'

  // Las incidencias no identificadas no restan días: quedan pendientes de
  // revisión en Validación.
  const incidenciasAplicables = (lista: IncidenciaImportada[]): IncidenciaImportada[] =>
    lista.filter((inc) => inc.estado !== 'no_identificada')

  const pendientesDeRevision = (lista: IncidenciaImportada[]): IncidenciaImportada[] =>
    lista.filter((inc) => inc.estado === 'no_identificada')

  const estadoValidacion = (trabajador: ImportedWorker): EstadoValidacion =>
    getObservaciones(trabajador).length === 0 ? 'Completo' : 'Pendiente'

  const ESTADO_VALIDACION_ESTILO: Record<EstadoValidacion, string> = {
    Completo: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30',
    Pendiente: 'bg-amber-500/10 text-amber-400 border-amber-500/30',
  }

  function getValidacionCalculada(trabajador: ImportedWorker, jornadas: Record<string, Jornada>, fechasEditadas: Record<string, { fechaInicio?: string; fechaCese?: string }>, mesesEditados: Record<string, string[]>, incidenciasEditadas: Record<string, IncidenciaImportada[]>): { posibles: number; noLaborados: number; efectivos: number; estado: EstadoValidacion; observaciones: string[] } {
    const jornada = (jornadas[trabajador.dni] ?? trabajador.jornada) as Jornada | undefined
    const fechas = fechasEditadas[trabajador.dni] ?? {}
    const incidencias = incidenciasEditadas[trabajador.dni] ?? trabajador.incidencias ?? []
    const noLaborados = calcularDiasNoLaborados(
      incidenciasAplicables(incidencias).map(inc => ({
        cantidadDias: inc.cantidadDias,
        diasNeto: inc.diasNeto,
        fechaInicio: parseIncFecha(inc.fechaInicio),
        fechaFin: parseIncFecha(inc.fechaFin),
      })),
      jornada ?? 5,
      new Set(trabajador.feriados ?? [])
    )
    const mesesConDias = new Set(mesesTrabajados(trabajador, mesesEditados[trabajador.dni]))
    const posibles = trabajador.diasPosibles
      ?? (trabajador.laborablesPorMes
        ? months.reduce((sum, mes) => sum + (mesesConDias.has(mes) ? (trabajador.laborablesPorMes?.[mes] ?? 0) : 0), 0)
        : (trabajador.diasTrabajados?.total ?? 0))
    const efectivos = Math.max(0, posibles - noLaborados)
    const ingreso = fechaIngresoTrabajador(trabajador, fechas.fechaInicio)

    const observaciones: string[] = []
    if (trabajador.remuneraciones.total <= 0) observaciones.push('Sin remuneraciones registradas para el ejercicio.')
    if (trabajador.remuneracionIncompleta) observaciones.push('Remuneración anual pendiente: el archivo no contiene los 12 meses necesarios para calcularla.')
    if (posibles <= 0) observaciones.push('Sin días laborados registrados para el ejercicio.')
    if (!ingreso.iso) observaciones.push('Fecha de ingreso pendiente de registrar.')
    if (!`${trabajador.apellidoPaterno}${trabajador.apellidoMaterno}${trabajador.nombres}`.trim()) observaciones.push('Nombres incompletos en la base importada.')
    if (!jornadas[trabajador.dni] && !trabajador.jornada) observaciones.push('Jornada no disponible en el archivo; no se asume automáticamente.')
    if (noLaborados > posibles) observaciones.push(`Los días no laborados (${number.format(noLaborados)}) superan los días posibles (${number.format(posibles)}); revisa las incidencias.`)
    if (trabajador.diasEfectivosReferencia && trabajador.diasEfectivosReferencia > 0 && trabajador.diasEfectivosReferencia !== efectivos) {
      observaciones.push(`Discrepancia de días: el sistema calcula ${number.format(efectivos)} efectivos (${number.format(posibles)} posibles − ${number.format(noLaborados)} no laborados), pero el cuadro del Excel indica ${number.format(trabajador.diasEfectivosReferencia)}.`)
    }
    observaciones.push(...observacionesRemuneracionImportada(trabajador.remuneracionResumen))

    const estado: EstadoValidacion = observaciones.length === 0 ? 'Completo' : 'Pendiente'
    return { posibles, noLaborados, efectivos, estado, observaciones }
  }

  type Celda = string | number

  function writeTableExcel(ws: ExcelJS.Worksheet, headers: string[], rows: Celda[][], opts: { moneyCols?: number[]; pctCols?: number[]; widths?: number[] } = {}) {
    const { moneyCols = [], pctCols = [], widths } = opts
    widths?.forEach((w, i) => { ws.getColumn(i + 1).width = w })
    const start = ws.lastRow ? ws.lastRow.number + 1 : 1
    const headerRow = ws.getRow(start)
    headers.forEach((h, i) => {
      const c = headerRow.getCell(i + 1)
      c.value = h
      c.font = { bold: true, color: { argb: 'FF4C1D95' } }
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F3FF' } }
      c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
    })
    let row = start + 1
    for (const data of rows) {
      data.forEach((v, j) => {
        const c = ws.getCell(row, j + 1)
        c.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
        if (moneyCols.includes(j)) {
          c.value = v as number
          c.numFmt = '"S/ " #,##0.00'
        } else if (pctCols.includes(j)) {
          c.value = v as number
          c.numFmt = '0.0000"%"'
        } else {
          c.value = v
        }
      })
      row++
    }
  }

  async function descargarInformacion(nombreArchivo: string, construir: (lib: typeof ExcelJS) => ExcelJS.Workbook): Promise<void> {
    const wb = construir(ExcelJS)
    const buffer = await wb.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = nombreArchivo
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  function BotonExportarInformacion({ onClick, disabled, label = 'Exportar información' }: { onClick: () => void; disabled?: boolean; label?: string }) {
    return (
      <Button variant="outline" size="sm" onClick={onClick} disabled={disabled} title="Exportar a Excel la información calculada">
        <Download className="mr-2 size-4" />{label}
      </Button>
    )
  }

  async function exportarCarga(trabajadores: ImportedWorker[], trabajadoresDias: ImportedWorker[]): Promise<void> {
    const totalRemAnual = trabajadores.reduce((s, t) => s + (t.remuneraciones?.total ?? 0), 0)
    const totalDiasReg = trabajadoresDias.reduce((s, t) => s + (t.diasTrabajados?.total ?? 0), 0)
    const totalIncidencias = trabajadoresDias.reduce((s, t) => s + (t.incidencias?.length ?? 0), 0)

    await descargarInformacion('carga-informacion-calculada.xlsx', (ExcelJS) => {
      const wb = new ExcelJS.Workbook()

      const resumen = wb.addWorksheet('Resumen')
      writeTableExcel(resumen, ['Concepto', 'Valor'], [
        ['Trabajadores con remuneraciones', trabajadores.length],
        ['Total remuneración anual computable', totalRemAnual],
        ['Trabajadores con días laborados', trabajadoresDias.length],
        ['Total días laborados registrados', totalDiasReg],
        ['Incidencias registradas', totalIncidencias],
      ], { moneyCols: [1], widths: [40, 24] })

      const rem = wb.addWorksheet('Remuneraciones calculadas')
      writeTableExcel(rem, ['DNI', 'Apellido paterno', 'Apellido materno', 'Nombres', ...months.map((m) => MONTH_LABELS[m]), 'Total anual'], trabajadores.map((t) => [
        t.dni, t.apellidoPaterno, t.apellidoMaterno, t.nombres,
        ...months.map((m) => t.remuneraciones?.[m] ?? 0),
        t.remuneraciones?.total ?? 0,
      ]), { moneyCols: months.map((_, i) => i + 4).concat(16), widths: [14, 20, 20, 24, ...months.map(() => 12), 16] })

      const dias = wb.addWorksheet('Días laborados calculados')
      writeTableExcel(dias, ['DNI', 'Apellido paterno', 'Apellido materno', 'Nombres', ...months.map((m) => MONTH_LABELS[m]), 'Total días', 'Incidencias'], trabajadoresDias.map((t) => [
        t.dni, t.apellidoPaterno, t.apellidoMaterno, t.nombres,
        ...months.map((m) => t.diasTrabajados?.[m] ?? 0),
        t.diasTrabajados?.total ?? 0,
        t.incidencias?.length ?? 0,
      ]), { widths: [14, 20, 20, 24, ...months.map(() => 12), 14, 12] })

      return wb
    })
  }

  async function exportarValidacion(trabajadores: ImportedWorker[], jornadas: Record<string, Jornada>, fechasEditadas: Record<string, { fechaInicio?: string; fechaCese?: string }>, mesesEditados: Record<string, string[]>, incidenciasEditadas: Record<string, IncidenciaImportada[]>): Promise<void> {
    const conRemuneraciones = trabajadores.filter((t) => t.remuneraciones.total > 0).length
    const conDias = trabajadores.filter((t) => (t.diasTrabajados?.total ?? 0) > 0).length
    const totalIncidencias = trabajadores.reduce((s, t) => s + (t.incidencias?.length ?? 0), 0)
    const completos = trabajadores.filter((t) => getValidacionCalculada(t, jornadas, fechasEditadas, mesesEditados, incidenciasEditadas).estado === 'Completo').length

    await descargarInformacion('validacion-informacion-calculada.xlsx', (ExcelJS) => {
      const wb = new ExcelJS.Workbook()

      const resumen = wb.addWorksheet('Resumen')
      writeTableExcel(resumen, ['Concepto', 'Valor'], [
        ['Registros', trabajadores.length],
        ['Con remuneraciones', conRemuneraciones],
        ['Con días laborados', conDias],
        ['Incidencias', totalIncidencias],
        ['Trabajadores completos', completos],
        ['Trabajadores pendientes', trabajadores.length - completos],
      ], { widths: [40, 24] })

      const ws = wb.addWorksheet('Validación calculada')
      const filas = trabajadores.map((t) => {
        const v = getValidacionCalculada(t, jornadas, fechasEditadas, mesesEditados, incidenciasEditadas)
        const jornada = (jornadas[t.dni] ?? t.jornada) as Jornada | undefined
        const periodos = getPeriodosTrabajados(t, mesesEditados[t.dni])
        return [
          t.dni,
          `${t.apellidoPaterno} ${t.apellidoMaterno}`.trim() || t.nombres || 'Sin nombre',
          jornada ? `${jornada} días/semana` : 'No disponible',
          fechaIngresoTrabajador(t, fechasEditadas[t.dni]?.fechaInicio).etiqueta,
          fechaCeseTrabajador(t, fechasEditadas[t.dni]?.fechaCese).etiqueta,
          periodos.filter((p) => p.estado === 'Trabajó' || p.estado === 'Reingresó').length,
          periodos.filter((p) => p.estado === 'No trabajó').length,
          v.posibles,
          v.noLaborados,
          v.efectivos,
          t.remuneraciones.total,
          pendientesDeRevision(incidenciasEditadas[t.dni] ?? t.incidencias ?? []).length,
          v.estado,
        ]
      })
      writeTableExcel(ws, ['DNI', 'Trabajador', 'Jornada', 'Ingreso', 'Cese', 'Meses trabajados', 'Meses sin trabajar', 'Días posibles', 'Días no laborados', 'Días efectivos', 'Remuneración anual', 'Incidencias pendientes', 'Estado'], filas, { moneyCols: [10], widths: [14, 32, 16, 14, 14, 16, 18, 14, 16, 14, 18, 20, 14] })

      const incidenciasWs = wb.addWorksheet('Incidencias')
      const filasIncidencias: Celda[][] = []
      for (const t of trabajadores) {
        const lista = incidenciasEditadas[t.dni] ?? t.incidencias ?? []
        for (const inc of lista) {
          filasIncidencias.push([
            t.dni,
            `${t.apellidoPaterno} ${t.apellidoMaterno}`.trim() || t.nombres || 'Sin nombre',
            inc.descripcion || inc.tipo || '',
            formatearPeriodo(inc.periodo) || SIN_DATO,
            inc.periodo?.anio ?? '',
            formatInputDate(inc.fechaInicio) || SIN_DATO,
            formatInputDate(inc.fechaFin) || SIN_DATO,
            inc.cantidadDias,
            inc.diasNeto ?? '',
            inc.estado === 'no_identificada' ? 'Pendiente de identificar' : 'Identificada',
            inc.origen?.map((o) => `${o.hoja}!${String.fromCharCode(64 + Math.min(o.columna, 26))}${o.fila}`).join(' · ') ?? '',
          ])
        }
      }
      writeTableExcel(incidenciasWs, ['DNI', 'Trabajador', 'Tipo', 'Período', 'Año del período', 'Fecha inicio', 'Fecha fin', 'Días', 'Días netos', 'Estado', 'Origen'], filasIncidencias, { widths: [14, 32, 18, 16, 18, 14, 14, 10, 12, 22, 26] })

      return wb
    })
  }

  async function exportarDistribucion(results: EmployeeUtilityResult[], totals: ReturnType<typeof getTotals>): Promise<void> {
    const totalDiasEfectivos = results.reduce((s, r) => s + (r.diasEfectivos ?? 0), 0)
    const totalRemuneraciones = results.reduce((s, r) => s + (r.remuneracionComputable ?? 0), 0)
    const fondoPorDias = totals.fund / 2
    const fondoPorRemuneraciones = totals.fund / 2

    await descargarInformacion('distribucion-informacion-calculada.xlsx', (ExcelJS) => {
      const wb = new ExcelJS.Workbook()

      const resumen = wb.addWorksheet('Resumen del fondo')
      writeTableExcel(resumen, ['Concepto', 'Valor'], [
        ['Fondo total', totals.fund],
        ['Fondo por días (50%)', fondoPorDias],
        ['Fondo por remuneraciones (50%)', fondoPorRemuneraciones],
        ['Total días efectivos', totalDiasEfectivos],
        ['Total remuneraciones', totalRemuneraciones],
        ['Trabajadores', results.length],
        ['A distribuir', totals.distributed],
        ['Quinta categoría total', totals.totalQuintaCategoria],
        ['Remanente', totals.remainder],
        ['Topes aplicados', totals.capped],
      ], { moneyCols: [1], widths: [40, 24] })

      const ws = wb.addWorksheet('Distribución calculada')
      const filas = results.map((r) => {
        const pctDias = totalDiasEfectivos > 0 ? (r.diasEfectivos / totalDiasEfectivos) * 100 : 0
        const pctRem = totalRemuneraciones > 0 ? (r.remuneracionComputable / totalRemuneraciones) * 100 : 0
        return [
          r.code,
          r.name,
          r.diasEfectivos,
          pctDias,
          r.distribucion.utilidadPorDias,
          r.remuneracionComputable,
          pctRem,
          r.distribucion.utilidadPorRemuneraciones,
          r.distribucion.utilidadBruta,
          r.capApplied ? r.cap : 0,
          r.remainder,
          r.distribucion.quintaCategoria,
          r.distribucion.utilidadNeta,
        ]
      })
      writeTableExcel(ws, ['DNI', 'Trabajador', 'Días efectivos', '% días', 'Importe días', 'Rem. computable', '% rem', 'Importe rem', 'Utilidad bruta', 'Tope', 'Excedente', 'Quinta categoría', 'Utilidad neta'], filas, { moneyCols: [4, 5, 7, 8, 9, 10, 11, 12], pctCols: [3, 6], widths: [14, 30, 14, 14, 16, 16, 14, 16, 16, 14, 14, 14, 14] })

      return wb
    })
  }

  function EstadoValidacionBadge({ estado }: { estado: EstadoValidacion }) {
    return <Badge variant="outline" className={`font-medium ${ESTADO_VALIDACION_ESTILO[estado]}`}>{estado}</Badge>
  }

  function DetalleSection({ titulo, descripcion, children }: { titulo: string; descripcion?: string; children: React.ReactNode }) {
    return (
      <div className="rounded-xl border border-border/60 bg-muted/10 p-4">
        <p className="text-sm font-semibold">{titulo}</p>
        {descripcion && <p className="mt-0.5 text-xs text-muted-foreground">{descripcion}</p>}
        <div className="mt-4">{children}</div>
      </div>
    )
  }

  function TablaVertical({ filas }: { filas: { label: string; valor: React.ReactNode }[] }) {
    return (
      <div className="overflow-hidden rounded-lg border border-border/60 bg-background">
        <table className="w-full border-collapse text-sm">
          <tbody>
            {filas.map(({ label, valor }) => (
              <tr key={label} className="border-b border-border/60 last:border-b-0">
                <td className="w-1/3 px-4 py-2 align-middle text-xs font-medium uppercase tracking-wide text-muted-foreground sm:w-1/2">{label}</td>
                <td className="px-4 py-2 font-semibold">{valor}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )
  }

  /** Resumen de períodos: se calcula con los mismos datos que la grilla. */
  function ResumenPeriodos({ periodos }: { periodos: { mes: string; nombre: string; estado: EstadoPeriodo }[] }) {
    return (
      <div className="mt-4 flex flex-wrap gap-2">
        <div className="rounded-lg border border-emerald-500/30 bg-emerald-500/10 px-3 py-1.5 text-xs font-semibold text-emerald-400">
          {periodos.filter((p) => p.estado === 'Trabajó').length} meses trabajados
        </div>

        <div className="rounded-lg border border-amber-500/30 bg-amber-500/10 px-3 py-1.5 text-xs font-semibold text-amber-400">
          {periodos.filter((p) => p.estado === 'Reingresó').length} reingresos
        </div>

        <div className="rounded-lg border border-slate-500/30 bg-slate-500/10 px-3 py-1.5 text-xs font-semibold text-slate-400">
          {periodos.filter((p) => p.estado === 'No trabajó').length} sin trabajar
        </div>
      </div>
    )
  }

  function PeriodosTrabajados({ worker }: { worker: ImportedWorker }) {
    const periodos = getPeriodosTrabajados(worker)

    return (
      <div className="mt-5 rounded-xl border border-border/60 bg-muted/10 p-4">
        <div className="mb-1 flex items-center justify-between">
          <div>
            <p className="text-sm font-semibold">
              Períodos trabajados
            </p>
            <p className="text-xs text-muted-foreground">
              Un mes cuenta como trabajado cuando el período tiene días registrados.
            </p>
          </div>
        </div>

        <div className="mt-4 grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6">
          {periodos.map(({ nombre, estado }) => (
            <div
              key={nombre}
              className={`rounded-lg border p-3 ${ESTADO_PERIODO_STYLE[estado]}`}
            >
              <p className="text-[11px] font-medium text-muted-foreground">
                {nombre}
              </p>
              <p className="mt-1 text-xs font-semibold">
                {estado}
              </p>
        </div>
        ))}
    </div>

        <ResumenPeriodos periodos={periodos} />
      </div>
    )
  }

  function PeriodosEditables({ trabajador, mesesEditados, setMesesEditados }: { trabajador: ImportedWorker; mesesEditados: Record<string, string[]>; setMesesEditados: React.Dispatch<React.SetStateAction<Record<string, string[]>>> }) {
    const actuales = mesesTrabajados(trabajador, mesesEditados[trabajador.dni])
    const periodos = getPeriodosTrabajados(trabajador, actuales)

    const alternar = (mes: string) => {
      setMesesEditados((prev) => {
        const actual = new Set(mesesTrabajados(trabajador, prev[trabajador.dni]))
        if (actual.has(mes)) actual.delete(mes)
        else actual.add(mes)
        return { ...prev, [trabajador.dni]: [...actual] }
      })
    }

    return (
      <div>
        <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6">
          {periodos.map(({ mes, nombre, estado }) => (
            <button
              key={nombre}
              type="button"
              onClick={() => alternar(mes)}
              className={`rounded-lg border p-3 text-left transition-colors hover:border-primary/50 ${ESTADO_PERIODO_STYLE[estado]}`}
            >
              <p className="text-[11px] font-medium text-muted-foreground">
                {nombre}
              </p>
              <p className="mt-1 text-xs font-semibold">
                {estado}
              </p>
            </button>
          ))}
        </div>

        <ResumenPeriodos periodos={periodos} />
      </div>
    )
  }

  function DetalleValidacion({ trabajador: t, jornadas, setJornadas, fechasEditadas, setFechasEditadas, mesesEditados, setMesesEditados, incidenciasEditadas, setIncidenciasEditadas, onReprocesar }: {
    trabajador: ImportedWorker
    jornadas: Record<string, Jornada>
    setJornadas: React.Dispatch<React.SetStateAction<Record<string, Jornada>>>
    fechasEditadas: Record<string, { fechaInicio?: string; fechaCese?: string }>
    setFechasEditadas: React.Dispatch<React.SetStateAction<Record<string, { fechaInicio?: string; fechaCese?: string }>>>
    mesesEditados: Record<string, string[]>
    setMesesEditados: React.Dispatch<React.SetStateAction<Record<string, string[]>>>
    incidenciasEditadas: Record<string, IncidenciaImportada[]>
    setIncidenciasEditadas: React.Dispatch<React.SetStateAction<Record<string, IncidenciaImportada[]>>>
    onReprocesar: () => void
  }) {
    const jornada = (jornadas[t.dni] ?? t.jornada) as Jornada | undefined
    const fechas = fechasEditadas[t.dni] ?? {}
    const incidencias = incidenciasEditadas[t.dni] ?? t.incidencias ?? []
    const noLaborados = calcularDiasNoLaborados(
      incidenciasAplicables(incidencias).map(inc => ({
        cantidadDias: inc.cantidadDias,
        diasNeto: inc.diasNeto,
        fechaInicio: parseIncFecha(inc.fechaInicio),
        fechaFin: parseIncFecha(inc.fechaFin),
      })),
      jornada ?? 5,
      new Set(t.feriados ?? [])
    )
    const mesesConDias = new Set(mesesTrabajados(t, mesesEditados[t.dni]))
    // DÍAS POSIBLES / LABORABLES: los días posibles del formato CITIKOLD (columna L)
    // tienen prioridad; en otros formatos se usan los días por mes cargados ("dias mes").
    const posibles = t.diasPosibles
      ?? (t.laborablesPorMes
        ? months.reduce((sum, mes) => sum + (mesesConDias.has(mes) ? (t.laborablesPorMes?.[mes] ?? 0) : 0), 0)
        : (t.diasTrabajados?.total ?? 0))
    // DÍAS EFECTIVOS = DÍAS POSIBLES − DÍAS NO LABORADOS (un solo descuento)
    const efectivos = Math.max(0, posibles - noLaborados)
    const ingreso = fechaIngresoTrabajador(t, fechas.fechaInicio)
    const cese = fechaCeseTrabajador(t, fechas.fechaCese)

    const observaciones: string[] = []
    if (t.remuneraciones.total <= 0) observaciones.push('Sin remuneraciones registradas para el ejercicio.')
    if (posibles <= 0) observaciones.push('Sin días laborados registrados para el ejercicio.')
    if (!ingreso.iso) observaciones.push('Fecha de ingreso pendiente de registrar.')
    if (!`${t.apellidoPaterno}${t.apellidoMaterno}${t.nombres}`.trim()) observaciones.push('Nombres incompletos en la base importada.')
    if (!jornadas[t.dni] && !t.jornada) observaciones.push('Jornada no disponible en el archivo; no se asume automáticamente. Verifica la jornada real antes de reprocesar.')
    if (noLaborados > posibles) observaciones.push(`Los días no laborados (${number.format(noLaborados)}) superan los días posibles (${number.format(posibles)}); revisa las incidencias.`)
    if (t.diasEfectivosReferencia && t.diasEfectivosReferencia > 0 && t.diasEfectivosReferencia !== efectivos) {
      observaciones.push(`Discrepancia de días: el sistema calcula ${number.format(efectivos)} efectivos (${number.format(posibles)} posibles − ${number.format(noLaborados)} no laborados), pero el cuadro del Excel indica ${number.format(t.diasEfectivosReferencia)}.`)
    }

    const estado: EstadoValidacion = observaciones.length === 0 ? 'Completo' : 'Pendiente'
    const nombreCompleto = `${t.apellidoPaterno} ${t.apellidoMaterno}`.trim() || t.nombres || 'Sin nombre'

    const actualizarIncidencia = (index: number, cambios: Partial<IncidenciaImportada>) => {
      setIncidenciasEditadas((prev) => {
        const actual = prev[t.dni] ?? t.incidencias ?? []
        return { ...prev, [t.dni]: actual.map((inc, i) => (i === index ? { ...inc, ...cambios } : inc)) }
      })
    }

    const agregarIncidencia = () => {
      setIncidenciasEditadas((prev) => {
        const actual = prev[t.dni] ?? t.incidencias ?? []
        return { ...prev, [t.dni]: [...actual, { dni: t.dni, codigo: '', descripcion: '', cantidadDias: 0, fechaInicio: null, fechaFin: null }] }
      })
    }

    const eliminarIncidencia = (index: number) => {
      setIncidenciasEditadas((prev) => {
        const actual = prev[t.dni] ?? t.incidencias ?? []
        return { ...prev, [t.dni]: actual.filter((_, i) => i !== index) }
      })
    }

    return (
      <div className="flex flex-col gap-5 p-5 sm:p-6">
        {/* Cabecera del trabajador */}
        <div className="flex items-center gap-3">
          <Avatar className="size-12">
            <AvatarFallback className="bg-primary/10 text-primary">
              {nombreCompleto.split(' ').map((n) => n[0]).filter(Boolean).slice(0, 2).join('').toUpperCase() || 'TR'}
            </AvatarFallback>
          </Avatar>
          <div className="min-w-0 flex-1">
            <p className="truncate text-xl font-semibold tracking-tight">{nombreCompleto}</p>
            <p className="text-sm text-muted-foreground">DNI {t.dni} · {t.nombres || '—'}</p>
          </div>
          <EstadoValidacionBadge estado={estado} />
        </div>

        {/* Datos del trabajador */}
        <DetalleSection titulo="Datos del trabajador" descripcion="Información general de la base importada.">
          <TablaVertical
            filas={[
              { label: 'Apellidos y nombres', valor: nombreCompleto },
              { label: 'DNI', valor: t.dni },
              { label: 'Código', valor: t.dni },
              {
                label: 'Ingreso',
                valor: <Input type="date" value={ingreso.iso} onChange={(e) => setFechasEditadas((prev) => ({ ...prev, [t.dni]: { ...prev[t.dni], fechaInicio: e.target.value || undefined } }))} />,
              },
              {
                label: 'Cese',
                valor: <Input type="date" value={cese.iso} onChange={(e) => setFechasEditadas((prev) => ({ ...prev, [t.dni]: { ...prev[t.dni], fechaCese: e.target.value || undefined } }))} />,
              },
              {
                label: 'Inicio (fecha normalizada)',
                valor: ingreso.iso ? ingreso.etiqueta : <span className="font-normal text-muted-foreground">{SIN_DATO}</span>,
              },
              {
                label: 'Cese (fecha normalizada)',
                valor: cese.iso ? cese.etiqueta : <span className="font-normal text-muted-foreground">{SIN_DATO}</span>,
              },
              { label: 'Estado', valor: <EstadoValidacionBadge estado={estado} /> },
            ]}
          />
        </DetalleSection>

        {/* Jornada laboral */}
        <DetalleSection titulo="Jornada laboral" descripcion="Corrige la jornada semanal; se aplica al reprocesar.">
          <TablaVertical
            filas={[
              {
                label: 'Jornada',
                valor: (
                  <select
                    value={jornada ?? ''}
                    onChange={(e) => setJornadas((prev) => ({ ...prev, [t.dni]: Number(e.target.value) as Jornada }))}
                    className="h-9 w-full rounded-lg border border-border bg-background px-2 text-sm font-semibold"
                  >
                    {jornada === undefined && <option value="" disabled>No disponible en el archivo</option>}
                    <option value={6}>6 días/semana</option>
                    <option value={5}>5 días/semana</option>
                    <option value={4}>4 días/semana</option>
                  </select>
                ),
              },
              { label: 'Turno', valor: <span className="font-normal text-muted-foreground">—</span> },
              { label: 'Días de descanso', valor: <span className="font-normal text-muted-foreground">—</span> },
            ]}
          />
        </DetalleSection>

        {/* Períodos trabajados */}
        <DetalleSection titulo="Períodos trabajados" descripcion="Haz clic en cada mes para alternar entre trabajado y no trabajado.">
          <PeriodosEditables trabajador={t} mesesEditados={mesesEditados} setMesesEditados={setMesesEditados} />
        </DetalleSection>

        {/* Días */}
        <DetalleSection titulo="Días" descripcion="Resumen de días posibles, no laborados y efectivos.">
          <TablaVertical
            filas={[
              { label: 'Días posibles', valor: number.format(posibles) },
              { label: 'Días no laborados', valor: <span className="text-rose-400">{number.format(noLaborados)}</span> },
              { label: 'Días efectivos', valor: <span className="text-primary">{number.format(efectivos)}</span> },
            ]}
          />
          <div className="mt-3 flex flex-wrap items-center gap-x-2 gap-y-1 rounded-lg border border-primary/20 bg-primary/5 px-3 py-2 text-xs text-muted-foreground">
            <span>Fórmula:</span>
            <span className="font-semibold text-foreground">{number.format(posibles)}</span>
            <span>−</span>
            <span className="font-semibold text-rose-400">{number.format(noLaborados)}</span>
            <span>=</span>
            <span className="font-semibold text-primary">{number.format(efectivos)}</span>
            <span>días efectivos</span>
          </div>
          {t.diasEfectivosReferencia && t.diasEfectivosReferencia > 0 && (
            <p className="mt-2 text-xs text-muted-foreground">
              Referencia del cuadro del Excel: {number.format(t.diasEfectivosReferencia)} días efectivos
              {' '}{t.diasEfectivosReferencia === efectivos ? '· coincide con el cálculo.' : '· NO coincide; revisar el detalle del trabajador.'}
            </p>
          )}
          {t.diasTrabajados && (
            <div className="mt-4 grid grid-cols-2 gap-2 sm:grid-cols-3 md:grid-cols-6">
              {months.map((mes) => (
                <div key={mes} className="rounded-lg border border-border/60 bg-background px-2 py-2 text-center">
                  <p className="text-[10px] font-medium text-muted-foreground">{MONTH_LABELS[mes]}</p>
                  <p className="text-sm font-semibold">{t.diasTrabajados?.[mes] ?? 0}</p>
                </div>
              ))}
            </div>
          )}
        </DetalleSection>

        {/* Remuneraciones */}
        <DetalleSection titulo="Remuneraciones" descripcion="Mensual registrada y total anual computable.">
          <div className="grid gap-3 sm:grid-cols-2">
            <TablaVertical filas={months.slice(0, 6).map((mes) => ({ label: MONTH_LABELS[mes], valor: money.format(t.remuneraciones[mes]) }))} />
            <TablaVertical filas={months.slice(6).map((mes) => ({ label: MONTH_LABELS[mes], valor: money.format(t.remuneraciones[mes]) }))} />
          </div>
          <div className="mt-3 flex items-center justify-between rounded-lg border border-primary/20 bg-primary/5 p-3">
            <span className="text-sm font-medium">Remuneración total anual</span>
            <span className="font-semibold text-primary">{money.format(t.remuneraciones.total)}</span>
          </div>
        </DetalleSection>

        {/* Incidencias */}
        <DetalleSection titulo="Incidencias" descripcion="Cada incidencia conserva el período, el tipo y los días que aporta el archivo. Registra, corrige o elimina las que restan días efectivos.">
          {incidencias.length === 0 ? (
            <p className="rounded-lg border border-border/60 bg-background px-4 py-6 text-center text-sm text-muted-foreground">No tiene incidencias registradas.</p>
          ) : (
            <div className="space-y-3">
              {incidencias.map((incidencia, index) => {
                const periodo = formatearPeriodo(incidencia.periodo)
                const rango = `${formatInputDate(incidencia.fechaInicio) || '—'} al ${formatInputDate(incidencia.fechaFin) || '—'}`
                const dias = `${incidencia.cantidadDias} ${incidencia.cantidadDias === 1 ? 'día' : 'días'}`
                const sinPeriodoNiFechas = !periodo && rango === '— al —'
                const pendiente = incidencia.estado === 'no_identificada'

                return (
                  <div key={`${incidencia.codigo}-${index}`} className="rounded-lg border border-border/60 bg-background p-3">
                    <div className="mb-3 flex flex-wrap items-center gap-2">
                      <Badge variant="outline" className="font-medium">
                        {incidencia.descripcion || incidencia.tipo || 'Incidencia sin tipo'}
                      </Badge>
                      <span className="text-sm font-semibold">
                        {periodo ? `${periodo} · ${dias}` : `${rango} · ${dias}`}
                      </span>
                      {pendiente && (
                        <Badge variant="outline" className="border-amber-500/30 bg-amber-500/10 text-amber-400">
                          Pendiente de identificar
                        </Badge>
                      )}
                      {sinPeriodoNiFechas && !pendiente && (
                        <Badge variant="outline" className="border-slate-500/30 bg-slate-500/10 text-slate-400">
                          Período no disponible
                        </Badge>
                      )}
                    </div>

                    {pendiente && incidencia.motivo && (
                      <p className="mb-3 rounded-md border border-amber-500/30 bg-amber-500/5 px-3 py-2 text-xs text-amber-500">
                        {incidencia.motivo}
                      </p>
                    )}

                    <div className="grid grid-cols-2 gap-3 sm:grid-cols-[1fr_1fr_1fr_auto]">
                      <div className="col-span-2 sm:col-span-1">
                        <p className="mb-1 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">Tipo</p>
                        <Input value={incidencia.descripcion} onChange={(e) => actualizarIncidencia(index, { descripcion: e.target.value })} placeholder="Descripción de la incidencia" />
                      </div>
                      <div>
                        <p className="mb-1 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">Fecha inicio</p>
                        <Input type="date" value={formatInputDate(incidencia.fechaInicio)} onChange={(e) => actualizarIncidencia(index, { fechaInicio: e.target.value || null })} />
                      </div>
                      <div>
                        <p className="mb-1 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">Fecha fin</p>
                        <Input type="date" value={formatInputDate(incidencia.fechaFin)} onChange={(e) => actualizarIncidencia(index, { fechaFin: e.target.value || null })} />
                      </div>
                      <div className="col-span-2 flex items-end gap-2 sm:col-span-1">
                        <div className="flex-1">
                          <p className="mb-1 text-[11px] font-medium uppercase tracking-wide text-muted-foreground">Días</p>
                          <Input type="number" min={0} value={incidencia.cantidadDias} onChange={(e) => actualizarIncidencia(index, { cantidadDias: Number(e.target.value) || 0 })} />
                        </div>
                        <Button variant="ghost" size="icon" onClick={() => eliminarIncidencia(index)} aria-label="Eliminar incidencia">
                          <Trash2 className="size-4" />
                        </Button>
                      </div>
                    </div>

                    {periodo && (
                      <p className="mt-2 text-[11px] text-muted-foreground">
                        Período en el archivo: <span className="font-medium text-foreground">{incidencia.periodo?.texto || periodo}</span>
                        {incidencia.periodo?.anio ? ` · Año ${incidencia.periodo.anio}` : ' · El archivo no indica el año'}
                      </p>
                    )}
                  </div>
                )
              })}
            </div>
          )}
          <Button variant="outline" size="sm" className="mt-3" onClick={agregarIncidencia}>
            <Plus className="mr-2 size-4" />Agregar incidencia
          </Button>
        </DetalleSection>

        {/* Observaciones */}
        <DetalleSection titulo="Observaciones y datos pendientes de validar" descripcion="Elementos que requieren atención antes del reproceso.">
          {observaciones.length === 0 ? (
            <div className="flex items-center gap-2 rounded-lg border border-emerald-500/30 bg-emerald-500/10 px-4 py-3 text-sm text-emerald-400">
              <Check className="size-4" />Sin observaciones. La información del trabajador está completa.
            </div>
          ) : (
            <div className="space-y-2">
              {observaciones.map((obs) => (
                <div key={obs} className="flex items-start gap-2 rounded-lg border border-amber-500/30 bg-amber-500/10 px-4 py-3 text-sm text-amber-500">
                  <AlertCircle className="mt-0.5 size-4 shrink-0" />{obs}
                </div>
              ))}
            </div>
          )}
        </DetalleSection>

        {/* Estado de validación */}
        <DetalleSection titulo="Estado de validación" descripcion="Resultado de la revisión de este trabajador.">
          <TablaVertical
            filas={[
              { label: 'Estado del trabajador', valor: <EstadoValidacionBadge estado={estado} /> },
              { label: 'Con remuneración', valor: t.remuneraciones.total > 0 ? 'Sí' : 'No' },
              { label: 'Con días', valor: posibles > 0 ? 'Sí' : 'No' },
              { label: 'Incidencias', valor: number.format(incidencias.length) },
            ]}
          />
        </DetalleSection>

        {/* Acciones */}
        <div className="sticky bottom-0 z-10 -mx-5 flex items-center justify-between gap-2 border-t border-border/60 bg-popover/95 px-5 py-3 backdrop-blur-sm sm:-mx-6 sm:px-6">
          <span className="mr-auto text-xs text-muted-foreground">Los cambios se aplicarán al reprocesar los días y resultados.</span>
          <Button onClick={onReprocesar}>
            <Play className="mr-2 size-4" />Reprocesar días y resultados
          </Button>
        </div>
      </div>
    )
  }

  const MESES_EN_CERO: MonthlyValues = { enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0, julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 0 }

  function PanelValidacion({ trabajadores, jornadas, setJornadas, fechasEditadas, setFechasEditadas, mesesEditados, setMesesEditados, incidenciasEditadas, setIncidenciasEditadas, onReprocesar }: {
    trabajadores: ImportedWorker[]
    jornadas: Record<string, Jornada>
    setJornadas: React.Dispatch<React.SetStateAction<Record<string, Jornada>>>
    fechasEditadas: Record<string, { fechaInicio?: string; fechaCese?: string }>
    setFechasEditadas: React.Dispatch<React.SetStateAction<Record<string, { fechaInicio?: string; fechaCese?: string }>>>
    mesesEditados: Record<string, string[]>
    setMesesEditados: React.Dispatch<React.SetStateAction<Record<string, string[]>>>
    incidenciasEditadas: Record<string, IncidenciaImportada[]>
    setIncidenciasEditadas: React.Dispatch<React.SetStateAction<Record<string, IncidenciaImportada[]>>>
    onReprocesar: () => void
  }) {
    const [busqueda, setBusqueda] = useState('')
    const [seleccionadoDni, setSeleccionadoDni] = useState<string | null>(null)

    const conRemuneracion = trabajadores.filter((t) => t.remuneraciones.total > 0).length
    const conDias = trabajadores.filter((t) => (t.diasTrabajados?.total ?? 0) > 0).length
    const totalIncidencias = trabajadores.reduce((total, t) => total + (t.incidencias?.length ?? 0), 0)
    const seleccionado = trabajadores.find((t) => t.dni === seleccionadoDni) ?? null

    const filtrados = trabajadores.filter((t) => {
      const texto = `${t.apellidoPaterno} ${t.apellidoMaterno} ${t.nombres} ${t.dni}`.toLowerCase()
      return texto.includes(busqueda.toLowerCase().trim())
    })

    return (
      <>
        <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
          <Kpi label="Registros" value={number.format(trabajadores.length)} hint="Base consolidada" icon={Users} tone="teal" />
          <Kpi label="Con remuneraciones" value={number.format(conRemuneracion)} hint="Del total identificado" icon={ClipboardCheck} />
          <Kpi label="Con días laborados" value={number.format(conDias)} hint="Del total identificado" icon={CalendarDays} tone="amber" />
          <Kpi label="Incidencias" value={number.format(totalIncidencias)} hint="Ajustan los días efectivos" icon={AlertCircle} tone="rose" />
        </div>

        <Card>
          <CardHeader>
            <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <CardTitle>Lista de validación</CardTitle>
                <CardDescription>Selecciona un trabajador para revisar y corregir su información.</CardDescription>
              </div>
              <div className="flex gap-2">
                <div className="relative">
                  <Search className="absolute left-3 top-2.5 size-4 text-muted-foreground" />
                  <Input value={busqueda} onChange={(e) => setBusqueda(e.target.value)} placeholder="Buscar trabajador" className="w-full pl-9 sm:w-64" />
                </div>
                <Button variant="outline" size="icon">
                  <Filter className="size-4" />
                </Button>
              </div>
            </div>
          </CardHeader>
          <CardContent>
            <div className="overflow-x-auto">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Trabajador</TableHead>
                    <TableHead>Jornada</TableHead>
                    <TableHead>Días posibles</TableHead>
                    <TableHead>Días no laborados</TableHead>
                    <TableHead>Días efectivos</TableHead>
                    <TableHead>Remuneración anual</TableHead>
                    <TableHead>Estado</TableHead>
                    <TableHead className="text-right">Acciones</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {filtrados.length === 0 && (
                    <TableRow>
                      <TableCell colSpan={8} className="py-10 text-center text-sm text-muted-foreground">Sin trabajadores registrados</TableCell>
                    </TableRow>
                  )}
                  {filtrados.map((t) => {
                    const jornada = jornadas[t.dni] ?? t.jornada
                    const incidencias = incidenciasEditadas[t.dni] ?? t.incidencias ?? []
                    const noLaborados = calcularDiasNoLaborados(
                      incidenciasAplicables(incidencias).map(inc => ({
                        cantidadDias: inc.cantidadDias,
                        diasNeto: inc.diasNeto,
                        fechaInicio: parseIncFecha(inc.fechaInicio),
                        fechaFin: parseIncFecha(inc.fechaFin),
                      })),
                      jornada ?? 5,
                      new Set(t.feriados ?? [])
                    )
                    const posibles = t.diasPosibles ?? (t.diasTrabajados?.total ?? 0)
                    const efectivos = Math.max(0, posibles - noLaborados)
                    const estado = estadoValidacion(t)

                    return (
                      <TableRow key={t.dni} onClick={() => setSeleccionadoDni(t.dni)} className="cursor-pointer">
                        <TableCell>
                          <p className="font-medium">{`${t.apellidoPaterno} ${t.apellidoMaterno}`.trim() || t.nombres || 'Sin nombre'}</p>
                          <p className="text-xs text-muted-foreground">{t.dni}</p>
                        </TableCell>
                        <TableCell>
                          {jornada
                            ? <Badge variant="outline">{jornada} días/semana</Badge>
                            : <Badge variant="outline" className="text-muted-foreground">No disponible</Badge>}
                        </TableCell>
                        <TableCell>{number.format(posibles)}</TableCell>
                        <TableCell>{number.format(noLaborados)}</TableCell>
                        <TableCell>{number.format(efectivos)}</TableCell>
                        <TableCell>{money.format(t.remuneraciones.total)}</TableCell>
                        <TableCell><EstadoValidacionBadge estado={estado} /></TableCell>
                        <TableCell className="text-right">
                          <ChevronRight className="ml-auto size-4 text-muted-foreground" />
                        </TableCell>
                      </TableRow>
                    )
                  })}
                </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>

        <Sheet open={!!seleccionado} onOpenChange={(open) => { if (!open) setSeleccionadoDni(null) }}>
          <SheetContent className="w-full overflow-y-auto">
            <SheetHeader>
              <SheetTitle>Validación del trabajador</SheetTitle>
              <SheetDescription>Revisa y corrige la información antes de reprocesar los días y resultados.</SheetDescription>
            </SheetHeader>
            {seleccionado && (
              <DetalleValidacion
                trabajador={seleccionado}
                jornadas={jornadas}
                setJornadas={setJornadas}
                fechasEditadas={fechasEditadas}
                setFechasEditadas={setFechasEditadas}
                mesesEditados={mesesEditados}
                setMesesEditados={setMesesEditados}
                incidenciasEditadas={incidenciasEditadas}
                setIncidenciasEditadas={setIncidenciasEditadas}
                onReprocesar={() => { onReprocesar(); setSeleccionadoDni(null) }}
              />
            )}
          </SheetContent>
        </Sheet>
      </>
    )
  }

  function ImportCard({
  kind,
  trabajadores,
  setTrabajadores,
  setTrabajadoresDias,
}: {
  kind: "remunerations" | "worked-days"
  trabajadores: ImportedWorker[]
  setTrabajadores: React.Dispatch<React.SetStateAction<ImportedWorker[]>>
  setTrabajadoresDias?: React.Dispatch<React.SetStateAction<ImportedWorker[]>>
}) {
    const isRemunerations = kind === 'remunerations'

    const title = isRemunerations
      ? 'Remuneraciones'
      : 'Días laborados'

    const description = isRemunerations
      ? 'Importa la información mensual o consolidada de remuneraciones computables.'
      : 'Carga los días trabajados, ausencias y licencias del ejercicio.'

    const inputRef = useRef<HTMLInputElement>(null)
    const [loading, setLoading] = useState(false)

  const [diasImportados, setDiasImportados] = useState(0)

    const [importedFile, setImportedFile] = useState<string | null>(null)
  const [trabajadorSeleccionado, setTrabajadorSeleccionado] = useState<number | null>(null)
  const [incidenciaTrabajador, setIncidenciaTrabajador] = useState<number | null>(null)
  const [busqueda, setBusqueda] = useState('')
  const [filtroEstado, setFiltroEstado] = useState('todos')
  const limpiarNombre = (valor: ExcelJS.CellValue): string => {
  if (!valor) return ''

  if (valor instanceof Date) return ''

  let texto = String(valor).trim()
  // Elimina cualquier fecha JavaScript y todo lo que venga después
  texto = texto.replace(
    /\s+(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+[A-Z][a-z]{2}\s+\d{1,2}\s+\d{4}.*$/gi,
    ''
  )

  return texto.trim()
}
  function detalleAuditoriaRemuneracion(resultado: { importados: ImportedWorkerMotor[]; modo?: string }): string {
    const resumenes = resultado.importados
      .map((trabajador) => trabajador.remuneracionResumen)
      .filter((resumen): resumen is ResumenRemuneracion => Boolean(resumen))

    if (resumenes.length === 0) return ''

    const lineas: string[] = []
    const totalMeses = resumenes.filter((resumen) => !resumen.mensualDisponible && resumen.totalDeclarado !== undefined).length
    const duplicados = resumenes.reduce((acc, resumen) => acc + resumen.duplicados, 0)
    const conflictos = resumenes.reduce((acc, resumen) => acc + resumen.conflictos.length, 0)
    const desconocidos = [...new Set(resumenes.flatMap((resumen) => resumen.conceptosDesconocidos))]
    const noRemunerativos = [...new Set(resumenes.flatMap((resumen) => resumen.conceptosExcluidos))]

    lineas.push(`\nDetalle de importación: ${resultado.modo === 'atomico' ? 'conceptos por período' : 'cuadro mensual'}`)
    if (totalMeses > 0) lineas.push(`• ${totalMeses} trabajador(es) solo con total anual.`)
    if (duplicados > 0) lineas.push(`• ${duplicados} registro(s) duplicado(s) descartado(s).`)
    if (conflictos > 0) lineas.push(`• ${conflictos} diferencia(s) entre total declarado y detalle.`)
    if (desconocidos.length > 0) lineas.push(`• Conceptos fuera del catálogo excluidos: ${desconocidos.join(', ')}.`)
    if (noRemunerativos.length > 0) lineas.push(`• Conceptos no remunerativos excluidos: ${noRemunerativos.join(', ')}.`)
    lineas.push('\nRevisa el detalle en Validación.')

    return lineas.join('\n')
  }

  const handleFile = async (file: File) => {
      setLoading(true)

      try {
        const workbook = new ExcelJS.Workbook()
        const buffer = await file.arrayBuffer()

        await workbook.xlsx.load(buffer)

        if (isRemunerations) {
          const resultado = extraerRemuneraciones(workbook, trabajadores, { loteId: file.name })
          if (resultado.importados.length > 0) {
            setTrabajadores(adaptarTrabajadoresMotor(resultado.trabajadores))
            setImportedFile(file.name)
            alert(`Excel leído correctamente.\n\nTrabajadores encontrados: ${resultado.importados.length}${detalleAuditoriaRemuneracion(resultado)}`)
            return
          }
        } else {
          const resultado = extraerDias(workbook, trabajadores)
          if (resultado.diasImportados > 0) {
            setTrabajadoresDias?.(adaptarTrabajadoresMotor(resultado.trabajadores))
            setDiasImportados(resultado.diasImportados)
            setImportedFile(file.name)
            return
          }
        }

        if (!isRemunerations) {
  const normalizeHeader = (value: ExcelJS.CellValue) => {
    const text =
      typeof value === 'object' &&
      value !== null &&
      'text' in value
        ? String(value.text ?? '')
        : String(value ?? '')

    const normalized = text
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim()
      .replace(/\s+/g, ' ')
      .toUpperCase()

    return normalized === 'SETIEMBRE' ? 'SEPTIEMBRE' : normalized
  }

  const normalizeDni = (value: ExcelJS.CellValue) => {
    const dni = String(value ?? '')
      .trim()
      .replace(/\.0$/, '')

    if (/^\d{9}$/.test(dni)) {
      return dni
    }

    return /^\d{7,8}$/.test(dni)
      ? dni.padStart(8, '0')
      : null
  }

  const toNumber = (value: ExcelJS.CellValue): number => {
    if (typeof value === 'number') {
      return value
    }

    if (
      typeof value === 'object' &&
      value !== null &&
      'result' in value
    ) {
      return toNumber(value.result as ExcelJS.CellValue)
    }

    const numero = Number(
      String(value ?? '').replace(',', '.')
    )

    return Number.isFinite(numero) ? numero : 0
  }
  // PASO 1: VERIFICAR LAS HOJAS DEL EXCEL
  const nombresHojas = workbook.worksheets.map(
    (sheet) => sheet.name
  )
  console.log('======================================')
  console.log('EXCEL DE DÍAS CARGADO')
  console.log('HOJAS ENCONTRADAS:')
  console.log(nombresHojas)
  console.log('======================================')
  // PASO 2: BUSCAR HOJA "dias mes"
  const worksheetDiasMes =
    workbook.getWorksheet('dias mes')

  const citikoldDias =
    workbook.getWorksheet('DÍAS EFECTIVOS LABORADOS 2025')

  const citikoldIncidencias =
    workbook.getWorksheet('DIAS NO LABORADOS')

  // Formato "DIAS + FALTAS": matriz mensual de días laborados y de faltas,
  // con los encabezados (DNI y meses) en la fila 1 de cada hoja.
  const worksheetDiasMatriz =
    workbook.getWorksheet('DIAS')

  const worksheetFaltas =
    workbook.getWorksheet('FALTAS')

  const detectarColumnasMensuales = (ws: ExcelJS.Worksheet): Map<string, number> => {
    const columnas = new Map<string, number>()
    ws.getRow(1).eachCell((cell, col) => {
      const nombre = normalizeHeader(cell.value)
      const mes = months.find((m) => m.toUpperCase() === nombre)
      if (mes) columnas.set(mes, col)
    })
    return columnas
  }

  const hayFormatoDiasFaltas =
    !!worksheetDiasMatriz &&
    detectarColumnasMensuales(worksheetDiasMatriz).size > 0

  if (
    !worksheetDiasMes &&
    !citikoldDias &&
    !citikoldIncidencias &&
    !hayFormatoDiasFaltas
  ) {
    throw new Error(
      `No se encontró la hoja "dias mes".

Hojas encontradas:
${nombresHojas.join(', ')}`
    )
  }

  if (
    !worksheetDiasMes &&
    (citikoldDias || citikoldIncidencias)
  ) {
    // ============================================================
    // FORMATO CITIKOLD: "DÍAS EFECTIVOS LABORADOS 2025" + "DIAS NO LABORADOS"
    // ============================================================
    console.log('FORMATO CITIKOLD DE DÍAS DETECTADO')

    const daysByDniCitikold = new Map<string, MonthlyValues>()
    const incidenciasPorDniCitikold = new Map<string, IncidenciaImportada[]>()
    const fechasPorDniCitikold = new Map<string, { fechaInicio?: Date; fechaCese?: Date }>()
    const laborablesPorDniCitikold = new Map<string, MonthlyValues>()
    const jornadaPorDniCitikold = new Map<string, Jornada>()
    const diasPosiblesPorDni = new Map<string, number>()
    const deduccionesPorDni = new Map<string, number>()
    const efectivosReferenciaPorDni = new Map<string, number>()
    // Feriados nacionales del formato CITIKOLD (rango R8:R23, columna 18)
    const feriadosSet = new Set<string>()
    const nombresPorDniCitikold = new Map<
      string,
      { apellidoPaterno: string; apellidoMaterno: string; nombres: string }
    >()

    if (citikoldDias) {
      console.log('HOJA "DÍAS EFECTIVOS LABORADOS 2025" ENCONTRADA')
      
      // Helper para convertir fechas Excel a Date
      const parseExcelDate = (value: ExcelJS.CellValue): Date | null => {
        if (!value) return null
        if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value
        if (typeof value === 'number') {
          const excelEpoch = Date.UTC(1899, 11, 30)
          return new Date(excelEpoch + value * 86_400_000)
        }
        if (typeof value === 'string') {
          const date = new Date(value)
          return Number.isNaN(date.getTime()) ? null : date
        }
        if (typeof value === 'object' && value !== null) {
          if ('result' in value) return parseExcelDate(value.result as ExcelJS.CellValue)
          if ('text' in value && typeof value.text === 'string') return parseExcelDate(value.text)
        }
        return null
      }

      // Feriados nacionales del formato CITIKOLD (rango R8:R23, columna 18)
      for (let fr = 8; fr <= 23; fr++) {
        const f = parseExcelDate(citikoldDias.getRow(fr).getCell(18).value)
        if (f) feriadosSet.add(f.toISOString().slice(0, 10))
      }

      // Días laborables reales (según jornada y feriados) por mes de 2025,
      // dentro del período [fechaInicio, fechaCese].
      const diasLaborablesBrutosPorMes = (
        fechaInicio: Date | null,
        fechaCese: Date | null,
        jornada: string
      ): number[] => {
        const mes: number[] = new Array(12).fill(0)
        const esLV = (jornada || 'L-V').toUpperCase().includes('S') === false

        const inicio =
          fechaInicio && fechaInicio.getFullYear() === 2025
            ? new Date(Date.UTC(2025, fechaInicio.getUTCMonth(), fechaInicio.getUTCDate()))
            : new Date(Date.UTC(2025, 0, 1))

        const fin =
          fechaCese && fechaCese.getFullYear() === 2025
            ? new Date(Date.UTC(2025, fechaCese.getUTCMonth(), fechaCese.getUTCDate()))
            : new Date(Date.UTC(2025, 11, 31))

        const cur = new Date(inicio)
        while (cur <= fin) {
          const dow = cur.getUTCDay()
          const laborable = esLV ? dow !== 0 && dow !== 6 : dow !== 0
          if (laborable && !feriadosSet.has(cur.toISOString().slice(0, 10))) {
            mes[cur.getUTCMonth()]++
          }
          cur.setUTCDate(cur.getUTCDate() + 1)
        }
        return mes
      }

      // Distribuye el total efectivo de forma proporcional a los días
      // laborables brutos de cada mes, asegurando que la suma de meses
      // sea EXACTAMENTE igual al total (método del mayor residuo).
      const distribuirTotalPorMes = (brutos: number[], total: number): MonthlyValues => {
        const out: MonthlyValues = { ...MESES_EN_CERO, total }
        const suma = brutos.reduce((a, b) => a + b, 0)
        if (suma <= 0) return out

        const racionales = brutos.map((b) => (total * b) / suma)
        const enteros = racionales.map(Math.floor)
        let resto = total - enteros.reduce((a, b) => a + b, 0)

        const fracciones = racionales
          .map((v, i) => [i, v - Math.floor(v)] as [number, number])
          .sort((a, b) => b[1] - a[1])

        for (let k = 0; k < resto; k++) {
          enteros[fracciones[k]?.[0] ?? 0]++
        }

        months.forEach((m, i) => {
          out[m] = enteros[i]
        })
        return out
      }

      citikoldDias.eachRow((row, rowNumber) => {
        if (rowNumber < 8) return

        const dni = normalizeDni(row.getCell(2).value)
        if (!dni) return

        const nombreCompleto = limpiarNombre(row.getCell(3).value)
        const partes = nombreCompleto.split(/\s+/).filter(Boolean)

        nombresPorDniCitikold.set(dni, {
          apellidoPaterno: partes[0] ?? '',
          apellidoMaterno: partes[1] ?? '',
          nombres: partes.slice(2).join(' ') || nombreCompleto,
        })

        // Formato CITIKOLD:
        //  - L (columna 12) = DÍAS POSIBLES (potenciales del periodo)
        //  - M (columna 13) = DEDUCCIONES (días no laborados a descontar)
        //  - O (columna 15) = DÍAS EFECTIVOS (neto ya calculado por el Excel: L − M)
        // El sistema calcula DÍAS EFECTIVOS = DÍAS POSIBLES − DÍAS NO LABORADOS
        // (un solo descuento) y usa O únicamente como referencia de validación.
        const diasPosibles =
          toNumber(row.getCell(12).value) ||
          toNumber(row.getCell(7).value)
        const diasNoLaboradosReferencia =
          toNumber(row.getCell(13).value)
        const diasEfectivosReferencia =
          toNumber(row.getCell(15).value) ||
          toNumber(row.getCell(7).value)

        // Fechas de ingreso/cese (columnas fijas del formato CITIKOLD)
        const fechaInicioRaw = row.getCell(4).value
        const fechaCeseRaw = row.getCell(6).value
        
        const fechaInicio = parseExcelDate(fechaInicioRaw)
        const fechaCese = parseExcelDate(fechaCeseRaw)

        // Jornada del trabajador (columna 16: "L-V" o "L-S")
        const jornada = String(row.getCell(16).value ?? 'L-V').trim()

        // Días laborables reales por mes distribuidos para sumar los días posibles
        const brutos = diasLaborablesBrutosPorMes(fechaInicio, fechaCese, jornada)
        const distribuido = distribuirTotalPorMes(brutos, diasPosibles)

        // Guardar información adicional para cálculo de períodos
        daysByDniCitikold.set(dni, distribuido)

        // Referencias del cuadro para validación y reporte
        diasPosiblesPorDni.set(dni, diasPosibles)
        deduccionesPorDni.set(dni, diasNoLaboradosReferencia)
        efectivosReferenciaPorDni.set(dni, diasEfectivosReferencia)

        // Guardar los días laborables brutos por mes del tramo trabajado
        const laborablesPorMes = { ...MESES_EN_CERO }
        months.forEach((m, i) => { laborablesPorMes[m] = brutos[i] })
        laborablesPorMes.total = brutos.reduce((a, b) => a + b, 0)
        laborablesPorDniCitikold.set(dni, laborablesPorMes)

        // Guardar la jornada semanal del trabajador (L-S = 6 días, L-V = 5)
        const jornadaTexto = String(jornada || 'L-V').trim().toUpperCase()
        jornadaPorDniCitikold.set(dni, jornadaTexto.includes('S') ? 6 : 5)
        
        // Guardar fechas si están disponibles (para cálculo de períodos)
        if (fechaInicio || fechaCese) {
          fechasPorDniCitikold.set(dni, {
            fechaInicio: fechaInicio ?? fechasPorDniCitikold.get(dni)?.fechaInicio,
            fechaCese: fechaCese ?? fechasPorDniCitikold.get(dni)?.fechaCese,
          })
        }
      })
    }

    if (citikoldIncidencias) {
      console.log('HOJA "DIAS NO LABORADOS" ENCONTRADA')
      citikoldIncidencias.eachRow((row, rowNumber) => {
        if (rowNumber < 6) return

        const dni = normalizeDni(row.getCell(3).value)
        if (!dni) return

        const apellidoPaterno = limpiarNombre(row.getCell(4).value)
        const apellidoMaterno = limpiarNombre(row.getCell(5).value)
        const nombres = limpiarNombre(row.getCell(6).value)
        const previos = nombresPorDniCitikold.get(dni)

        nombresPorDniCitikold.set(dni, {
          apellidoPaterno:
            apellidoPaterno || previos?.apellidoPaterno || '',
          apellidoMaterno:
            apellidoMaterno || previos?.apellidoMaterno || '',
          nombres: nombres || previos?.nombres || '',
        })

        const descripcion = String(
          row.getCell(8).value ?? ''
        ).trim()
        const cantidadDias = toNumber(row.getCell(11).value)

        // El MES (columna 1) de la hoja de incidencias es el período al que
        // pertenece cada incidencia: se conserva en la incidencia.
        const mesCrudo = String(row.getCell(1).value ?? '').trim().toUpperCase()
        const mesNormalizado = mesCrudo === 'SETIEMBRE' ? 'SEPTIEMBRE' : mesCrudo
        const mesEncontrado = months.find(m => m.toUpperCase() === mesNormalizado)

        if (
          descripcion === '[object Object]' ||
          descripcion === 'NaN' ||
          (!descripcion && cantidadDias === 0)
        ) {
          return
        }

        // El período del archivo se conserva en la incidencia: mes de la hoja y
        // año de las fechas del trabajador (o el del ejercicio, si el archivo
        // lo aporta). Si no hay mes identificable, el período queda sin dato.
        const anioPeriodo = fechasPorDniCitikold.get(dni)?.fechaInicio?.getUTCFullYear() ?? null
        const periodoCitikold: PeriodoIncidencia | null = mesEncontrado
          ? {
            texto: MONTH_LABELS[mesEncontrado],
            anio: anioPeriodo,
            mes: months.indexOf(mesEncontrado) + 1,
          }
          : leerPeriodo(row.getCell(1).value)

        const existentes =
          incidenciasPorDniCitikold.get(dni) ?? []
        existentes.push({
          dni,
          codigo: String(row.getCell(7).value ?? '').trim(),
          descripcion,
          cantidadDias,
          diasNeto: toNumber(row.getCell(14).value) || undefined,
          fechaInicio: row.getCell(15).value,
          fechaFin: row.getCell(16).value,
          periodo: periodoCitikold,
        })
        incidenciasPorDniCitikold.set(dni, existentes)
      })
    }

    setDiasImportados(daysByDniCitikold.size)

    const noEncontrados = [...daysByDniCitikold.keys()].filter(
      (dni) =>
        !trabajadores.some(
          (trabajador) =>
            String(trabajador.dni).trim().padStart(8, '0') === dni
        )
    )

    const coincidentes =
      daysByDniCitikold.size - noEncontrados.length
    console.log('[CITIKOLD DÍAS] Coincidentes con REM:', coincidentes)

    setTrabajadores((current) => {
      // Función compartida para calcular meses presentes en formato CITIKOLD
      // Reglas:
      // 1. Los meses presentes se derivan de los DÍAS POR MES del trabajador:
      //    un mes cuenta como presente si tiene días > 0. Esto refleja de forma
      //    natural los reingresos (meses con 0 días en medio de un período con
      //    días posteriores) y los ingresos/ceses a mitad de año.
      // 2. Si no hay desglose mensual (meses todos en cero), se usa la fecha
      //    de inicio/cese como fallback.
      const calcularPresentes = (dni: string, trabajador?: ImportedWorker): string[] => {
        const diasTrabajados = daysByDniCitikold.get(dni)
        const fechas = fechasPorDniCitikold.get(dni)
        const total = diasTrabajados?.total ?? 0

        if (total <= 0) return trabajador?.mesesPresentes ?? []

        // Si el desglose mensual tiene al menos un mes con días, derivar
        // los meses presentes directamente de los días por mes.
        if (diasTrabajados) {
          const conDias = months.filter((mes) => (diasTrabajados[mes] ?? 0) > 0)
          if (conDias.length > 0) {
            return conDias
          }
        }

        const fechaInicio = fechas?.fechaInicio
        const fechaCese = fechas?.fechaCese

        // Caso 1: Fecha de inicio conocida y corresponde al ejercicio 2025
        if (fechaInicio && fechaInicio.getFullYear() === 2025) {
          const mesInicio = fechaInicio.getMonth()
          const mesFin = fechaCese && fechaCese.getFullYear() === 2025
            ? fechaCese.getMonth()
            : 11
          return months.slice(mesInicio, mesFin + 1)
        }

        // Caso 2: Fecha de inicio anterior a 2025 o en enero 2025
        // Trabajador con vínculo previo: se asume año completo salvo que
        // los días totales sean insuficientes (posible cese/reingreso).
        if ((!fechaInicio || fechaInicio.getFullYear() < 2025) ) {
          if (fechaCese && fechaCese.getFullYear() === 2025) {
            const mesCese = fechaCese.getMonth()
            return months.slice(0, mesCese + 1)
          }

          const mesesAproximados = Math.min(12, Math.ceil(total / 20))
          if (mesesAproximados >= 10) {
            return [...months]
          }
          return months.slice(0, mesesAproximados)
        }

        // Caso 3: Sin fecha de inicio fiable
        const mesesAproximados = Math.min(12, Math.ceil(total / 20))
        return months.slice(0, mesesAproximados)
      }

      const actualizados = current.map((trabajador) => {
        const clave = String(trabajador.dni).trim().padStart(8, '0')
        const diasTrabajados = daysByDniCitikold.get(clave)
        const incidencias =
          incidenciasPorDniCitikold.get(clave) ?? []
        const infoNombres = nombresPorDniCitikold.get(clave)

        // Calcular mesesPresentes para formato Citikold
        let mesesPresentes = trabajador.mesesPresentes ?? []
        if (diasTrabajados && diasTrabajados.total > 0 && mesesPresentes.length === 0) {
          mesesPresentes = calcularPresentes(clave, trabajador)
        }

        // Obtener fechas guardadas
        const fechas = fechasPorDniCitikold.get(clave)

        return {
          ...trabajador,
          apellidoPaterno:
            trabajador.apellidoPaterno || infoNombres?.apellidoPaterno || '',
          apellidoMaterno:
            trabajador.apellidoMaterno || infoNombres?.apellidoMaterno || '',
          nombres: trabajador.nombres || infoNombres?.nombres || '',
          fechaInicio: fechas?.fechaInicio ?? trabajador.fechaInicio,
          fechaCese: fechas?.fechaCese ?? trabajador.fechaCese,
          ...(diasTrabajados ? { diasTrabajados } : {}),
          incidencias,
          mesesPresentes,
          jornada: jornadaPorDniCitikold.get(clave) ?? trabajador.jornada ?? 5,
          diasPosibles: diasPosiblesPorDni.get(clave) ?? trabajador.diasPosibles,
          diasNoLaboradosReferencia: deduccionesPorDni.get(clave) ?? trabajador.diasNoLaboradosReferencia,
          diasEfectivosReferencia: efectivosReferenciaPorDni.get(clave) ?? trabajador.diasEfectivosReferencia,
          feriados: [...feriadosSet],
          laborablesPorMes: laborablesPorDniCitikold.get(clave) ?? trabajador.laborablesPorMes,
        }
      })

      const nuevos = [...daysByDniCitikold.keys()]
        .filter(
          (dni) =>
            !current.some(
              (trabajador) =>
                String(trabajador.dni).trim().padStart(8, '0') === dni
            )
        )
        .map((dni) => {
          const infoNombres = nombresPorDniCitikold.get(dni)
          const diasTrabajados = daysByDniCitikold.get(dni)!
          const incidencias =
            incidenciasPorDniCitikold.get(dni) ?? []
          
          // Para formato Citikold, calcular mesesPresentes de forma robusta
          let mesesDias = calcularPresentes(dni)

          // Obtener fechas guardadas si están disponibles
          const fechas = fechasPorDniCitikold.get(dni)

          return {
            dni,
            apellidoPaterno: infoNombres?.apellidoPaterno ?? '',
            apellidoMaterno: infoNombres?.apellidoMaterno ?? '',
            nombres: infoNombres?.nombres ?? '',
            fechaInicio: fechas?.fechaInicio ?? null,
            fechaCese: fechas?.fechaCese ?? null,
            remuneraciones: MESES_EN_CERO,
            diasTrabajados,
            incidencias,
            mesesDias,
            mesesPresentes: mesesDias,
            jornada: jornadaPorDniCitikold.get(dni) ?? 5,
            diasPosibles: diasPosiblesPorDni.get(dni),
            diasNoLaboradosReferencia: deduccionesPorDni.get(dni),
            diasEfectivosReferencia: efectivosReferenciaPorDni.get(dni),
            feriados: [...feriadosSet],
            laborablesPorMes: laborablesPorDniCitikold.get(dni),
          }
        })

      return [...actualizados, ...nuevos]
    })

    setImportedFile(file.name)

    console.log('=== EXTRACCIÓN DE DÍAS (CITIKOLD) TERMINADA ===')
    console.log(
      'Días registrados:',
      daysByDniCitikold.size,
      'Incidencias:',
      incidenciasPorDniCitikold.size
    )

    return
  }

  if (!worksheetDiasMes && hayFormatoDiasFaltas) {
    // ============================================================
    // FORMATO DIAS + FALTAS: matriz mensual de días laborados (hoja
    // "DIAS") y de faltas / días no laborados (hoja "FALTAS"), con
    // los encabezados en la fila 1.
    // ============================================================
    console.log('FORMATO DIAS + FALTAS DE DÍAS DETECTADO')

    const anioEjercicioArchivo = inferirAnioEjercicio(workbook)

    const columnasDias = detectarColumnasMensuales(worksheetDiasMatriz!)
    const columnasFaltas = worksheetFaltas ? detectarColumnasMensuales(worksheetFaltas) : new Map<string, number>()

    const leerColumnaEncabezado = (ws: ExcelJS.Worksheet, nombres: string[]): number | null => {
      let encontrada: number | null = null
      ws.getRow(1).eachCell((cell, col) => {
        const nombre = normalizeHeader(cell.value)
        if (!encontrada && nombres.includes(nombre)) encontrada = col
      })
      return encontrada
    }

    const colDniDias = leerColumnaEncabezado(worksheetDiasMatriz!, ['DNI']) ?? 1
    const colApDias = leerColumnaEncabezado(worksheetDiasMatriz!, ['APELLIDO PATERNO'])
    const colAmDias = leerColumnaEncabezado(worksheetDiasMatriz!, ['APELLIDO MATERNO'])
    const colNombresDias = leerColumnaEncabezado(worksheetDiasMatriz!, ['NOMBRES'])

    const diasPorDni = new Map<string, { values: MonthlyValues; apellidoPaterno: string; apellidoMaterno: string; nombres: string }>()

    worksheetDiasMatriz!.eachRow((row, rowNumber) => {
      if (rowNumber < 2) return
      const dni = normalizeDni(row.getCell(colDniDias).value)
      if (!dni) return

      const values = { ...MESES_EN_CERO }
      for (const [mes, col] of columnasDias) {
        values[mes as keyof MonthlyValues] = toNumber(row.getCell(col).value)
      }
      values.total = months.reduce((sum, mes) => sum + values[mes], 0)

      diasPorDni.set(dni, {
        values,
        apellidoPaterno: colApDias ? limpiarNombre(row.getCell(colApDias).value) : '',
        apellidoMaterno: colAmDias ? limpiarNombre(row.getCell(colAmDias).value) : '',
        nombres: colNombresDias ? limpiarNombre(row.getCell(colNombresDias).value) : '',
      })
    })

    const incidenciasPorDni = new Map<string, IncidenciaImportada[]>()
    const nombresDesdeFaltas = new Map<string, { apellidoPaterno: string; apellidoMaterno: string; nombres: string }>()

    if (worksheetFaltas && columnasFaltas.size > 0) {
      const colDniFaltas = leerColumnaEncabezado(worksheetFaltas, ['DNI']) ?? 1
      const colApFaltas = leerColumnaEncabezado(worksheetFaltas, ['APELLIDO PATERNO'])
      const colAmFaltas = leerColumnaEncabezado(worksheetFaltas, ['APELLIDO MATERNO'])
      const colNombresFaltas = leerColumnaEncabezado(worksheetFaltas, ['NOMBRES'])

      worksheetFaltas.eachRow((row, rowNumber) => {
        if (rowNumber < 2) return
        const dni = normalizeDni(row.getCell(colDniFaltas).value)
        if (!dni) return

        nombresDesdeFaltas.set(dni, {
          apellidoPaterno: colApFaltas ? limpiarNombre(row.getCell(colApFaltas).value) : '',
          apellidoMaterno: colAmFaltas ? limpiarNombre(row.getCell(colAmFaltas).value) : '',
          nombres: colNombresFaltas ? limpiarNombre(row.getCell(colNombresFaltas).value) : '',
        })

        const incidencias: IncidenciaImportada[] = []
        for (const [mes, col] of columnasFaltas) {
          const cantidadDias = toNumber(row.getCell(col).value)
          if (cantidadDias > 0) {
            const nombreMes = mes as (typeof months)[number]
            incidencias.push({
              dni,
              codigo: '',
              descripcion: 'Faltas',
              cantidadDias,
              fechaInicio: null,
              fechaFin: null,
              // La columna del mes es el período de la incidencia.
              periodo: {
                texto: MONTH_LABELS[nombreMes],
                anio: anioEjercicioArchivo,
                mes: months.indexOf(nombreMes) + 1,
              },
            })
          }
        }
        if (incidencias.length > 0) incidenciasPorDni.set(dni, incidencias)
      })
    }

    setDiasImportados(diasPorDni.size)

    setTrabajadores((current) => {
      const actualizados = current.map((trabajador) => {
        const clave = String(trabajador.dni).trim().padStart(8, '0')
        const info = diasPorDni.get(clave)
        if (!info) return trabajador

        const nombresFaltas = nombresDesdeFaltas.get(clave)
        const incidencias = incidenciasPorDni.get(clave) ?? []
        const mesesDias = months.filter((mes) => (info.values[mes] ?? 0) > 0)

        return {
          ...trabajador,
          apellidoPaterno: trabajador.apellidoPaterno || info.apellidoPaterno || nombresFaltas?.apellidoPaterno || '',
          apellidoMaterno: trabajador.apellidoMaterno || info.apellidoMaterno || nombresFaltas?.apellidoMaterno || '',
          nombres: trabajador.nombres || info.nombres || nombresFaltas?.nombres || '',
          diasTrabajados: info.values,
          diasPosibles: info.values.total,
          incidencias,
          mesesDias,
          mesesPresentes: mesesDias,
        }
      })

      const nuevos = [...diasPorDni.keys()]
        .filter((dni) => !current.some((t) => String(t.dni).trim().padStart(8, '0') === dni))
        .map((dni) => {
          const info = diasPorDni.get(dni)!
          const nombresFaltas = nombresDesdeFaltas.get(dni)
          const mesesDias = months.filter((mes) => (info.values[mes] ?? 0) > 0)
          return {
            dni,
            apellidoPaterno: info.apellidoPaterno || nombresFaltas?.apellidoPaterno || '',
            apellidoMaterno: info.apellidoMaterno || nombresFaltas?.apellidoMaterno || '',
            nombres: info.nombres || nombresFaltas?.nombres || '',
            fechaInicio: null,
            fechaCese: null,
            remuneraciones: MESES_EN_CERO,
            diasTrabajados: info.values,
            diasPosibles: info.values.total,
            incidencias: incidenciasPorDni.get(dni) ?? [],
            mesesDias,
            mesesPresentes: mesesDias,
          }
        })

      return [...actualizados, ...nuevos]
    })

    setImportedFile(file.name)

    console.log('=== EXTRACCIÓN DE DÍAS (DIAS + FALTAS) TERMINADA ===')
    console.log('Días registrados:', diasPorDni.size, 'Incidencias:', incidenciasPorDni.size)

    return
  }

  console.log('HOJA "dias mes" ENCONTRADA')

  // Si llegamos aquí sin "dias mes", el formato CITIKOLD ya fue
  // procesado en el bloque anterior.
  if (!worksheetDiasMes) {
    return
  }
  // PASO 3: BUSCAR HOJA "DIAS"
  const worksheetDias =
    workbook.getWorksheet('DIAS')

  if (worksheetDias) {
    console.log('HOJA "DIAS" ENCONTRADA')
  } else {
    console.warn(
      'ADVERTENCIA: No se encontró la hoja "DIAS".'
    )
  }
  // PASO 4: MOSTRAR LAS COLUMNAS DE "dias mes"
  console.log('--------------------------------------')
  console.log('COLUMNAS DE "dias mes"')
  console.log('--------------------------------------')

  const encabezadosDiasMes: {
    columna: number
    original: string
    normalizado: string
  }[] = []

  worksheetDiasMes.getRow(5).eachCell(
    (cell, column) => {
      const original = String(
        cell.value ?? ''
      ).trim()

      const normalizado =
        normalizeHeader(cell.value)

      if (original) {
        encabezadosDiasMes.push({
          columna: column,
          original,
          normalizado,
        })

        console.log(
          `Columna ${column}: "${original}" → "${normalizado}"`
        )
      }
    }
  )
  // PASO 5: VERIFICAR COLUMNAS NECESARIAS
  const columnMap = new Map<string, number>()

  encabezadosDiasMes.forEach(
    ({ columna, normalizado }) => {
      if (normalizado) {
        columnMap.set(
          normalizado,
          columna
        )
      }
    }
  )

  const requiredColumns = [
    'DNI',
    ...months.map((month) =>
      month.toUpperCase()
    ),
    'TOTAL',
  ]

  const missingColumns =
    requiredColumns.filter(
      (column) =>
        !columnMap.has(column)
    )

  if (missingColumns.length > 0) {
    console.error(
      'COLUMNAS FALTANTES EN "dias mes":',
      missingColumns
    )

    throw new Error(
      `La hoja "dias mes" no tiene las columnas requeridas:

${missingColumns.join(', ')}`
    )
  }

  console.log(
    'COLUMNAS NECESARIAS ENCONTRADAS CORRECTAMENTE'
  )
  // PASO 6: EXTRAER LOS DÍAS DE "dias mes"
  const dniColumn =
    columnMap.get('DNI')!

  const daysByDni =
    new Map<string, MonthlyValues>()

  let processed = 0

  worksheetDiasMes.eachRow(
    (row, rowNumber) => {
      if (rowNumber < 6) {
        return
      }

      const dni =
        normalizeDni(
          row.getCell(
            dniColumn
          ).value
        )

      if (!dni) {
        return
      }

      processed += 1

      const monthValues =
        Object.fromEntries(
          months.map(
            (month) => [
              month,
              toNumber(
                row.getCell(
                  columnMap.get(
                    month.toUpperCase()
                  )!
                ).value
              ),
            ]
          )
        ) as Omit<
          MonthlyValues,
          'total'
        >

      const calculatedTotal =
        months.reduce(
          (sum, month) =>
            sum +
            monthValues[month],
          0
        )

      const total =
        columnMap.has('TOTAL')
          ? toNumber(
              row.getCell(
                columnMap.get(
                  'TOTAL'
                )!
              ).value
            )
          : calculatedTotal

      daysByDni.set(
        dni,
        {
          ...monthValues,
          total,
        }
      )
    }
  )

  console.log(
    'TRABAJADORES CON DÍAS EXTRAÍDOS:',
    daysByDni.size
  )

  const nombresPorDni =
    new Map<
      string,
      {
        apellidoPaterno: string
        apellidoMaterno: string
        nombres: string
      }
    >()
// PASO 7: EXTRAER INCIDENCIAS Y FERIADOS DE "DIAS"
const incidenciasPorDni =
  new Map<string, IncidenciaImportada[]>()

const feriadosImportados: FeriadoImportado[] = []
if (worksheetDias) {
  console.log('======================================')
  console.log('ANALIZANDO HOJA "DIAS"')
  console.log('======================================')
// PASO 7: EXTRAER INCIDENCIAS DE "DIAS"
// La hoja "DIAS" es un reporte PLAME.
// Los encabezados están distribuidos en las filas 9 y 10.
// Los datos comienzan en la fila 11.
//
// Estructura detectada:
// A = Mes
// B = DNI
// C = Apellido paterno
// D = Apellido materno
// E = Nombres
// F = Código
// G-I = Descripción
// J = Num. Días
// K = Otro dato del reporte
// L = Fecha inicial
// M = Fecha final
// N = Otro dato del reporte
  console.log('======================================')
  console.log('EXTRAYENDO INCIDENCIAS DE "DIAS"')
  console.log('======================================')
  const FILA_DATOS = 11
  const COL_DNI = 2
  const COL_CODIGO = 6
  const COL_DESCRIPCION = 7
  const COL_CANTIDAD_DIAS = 10
  const COL_FECHA_INICIO = 13
  const COL_FECHA_FIN = 14
  let totalIncidencias = 0
  worksheetDias.eachRow(
    (row, rowNumber) => {
      if (rowNumber < FILA_DATOS) {
        return
      }
      // DNI
      const dni =
        normalizeDni(
          row.getCell(
            COL_DNI
          ).value
        )
      if (!dni) {
        return
      }
      // Nombres (columnas C / D / E)
      const apellidoPaterno =
  limpiarNombre(row.getCell(3).value)

const apellidoMaterno =
  limpiarNombre(row.getCell(4).value)

const nombres =
  limpiarNombre(row.getCell(5).value)
      const nombresPrevios =
        nombresPorDni.get(dni)
      nombresPorDni.set(
        dni,
        {
          apellidoPaterno:
            apellidoPaterno ||
            nombresPrevios?.apellidoPaterno ||
            '',
          apellidoMaterno:
            apellidoMaterno ||
            nombresPrevios?.apellidoMaterno ||
            '',
          nombres:
            nombres ||
            nombresPrevios?.nombres ||
            '',
        }
      )
      // CÓDIGO
      const codigo =
        String(
          row.getCell(
            COL_CODIGO
          ).value ?? ''
        ).trim()
      // DESCRIPCIÓN
      const descripcion =
        String(
          row.getCell(
            COL_DESCRIPCION
          ).value ?? ''
        ).trim()
      // CANTIDAD DE DÍAS
      const cantidadDias =
        toNumber(
          row.getCell(
            COL_CANTIDAD_DIAS
          ).value
        )
      // FECHA INICIAL
      const fechaInicio =
        row.getCell(
          COL_FECHA_INICIO
        ).value
      // FECHA FINAL
      const fechaFin =
        row.getCell(
          COL_FECHA_FIN
        ).value
      // Ignorar filas sin incidencia o con datos basura
      // (celdas de fórmulas/formato al final de la columna)
      if (
        descripcion === '[object Object]' ||
        descripcion === 'NaN' ||
        (!descripcion &&
          cantidadDias === 0 &&
          !fechaInicio &&
          !fechaFin)
      ) {
        return
      }
      const incidencia: IncidenciaImportada =
        {
          dni,
          codigo,
          descripcion,
          cantidadDias,
          diasNeto:
            toNumber(
              row.getCell(12).value
            ) || undefined,
          fechaInicio,
          fechaFin,
        }
      const incidenciasExistentes =
        incidenciasPorDni.get(
          dni
        ) ?? []
      incidenciasExistentes.push(
        incidencia
      )
      const claveDni = String(dni)
  .trim()
  .padStart(8, '0')

incidenciasPorDni.set(claveDni, incidenciasExistentes)
      totalIncidencias += 1
    }
  )
  console.log(
    'INCIDENCIAS EXTRAÍDAS:',
    totalIncidencias
  )
  console.log(
    'TRABAJADORES CON INCIDENCIAS:',
    incidenciasPorDni.size
  )
  // ----------------------------------------------
  // Mostrar las primeras incidencias
  // para comprobar que están bien extraídas
  // ----------------------------------------------
  let contador = 0
  for (
    const [
      dni,
      incidencias,
    ] of incidenciasPorDni
  ) {
    console.log(
      `DNI ${dni}:`,
      incidencias
    )
    contador += 1
    if (contador >= 10) {
      break
    }
  }
  // ==========================================================
  // BUSCAR FERIADOS
  // ==========================================================
  console.log(
    '======================================'
  )
  console.log(
    'BUSCANDO FERIADOS EN "DIAS"...'
  )
  let feriadosHeaderRow: number | null =
    null
  let feriadosHeaderMap =
    new Map<string, number>()
  worksheetDias.eachRow(
    (row, rowNumber) => {
      if (
        feriadosHeaderRow !== null
      ) {
        return
      }
      const posiblesHeaders =
        new Map<string, number>()
      row.eachCell(
        (cell, column) => {
          const header =
            normalizeHeader(
              cell.value
            )
          if (header) {
            posiblesHeaders.set(
              header,
              column
            )
          }
        }
      )
      const tieneFecha =
        posiblesHeaders.has(
          'FECHA'
        )
      const tieneDescripcion =
        posiblesHeaders.has(
          'DESCRIPCION'
        ) ||
        posiblesHeaders.has(
          'FERIADO'
        )
      if (
        tieneFecha &&
        tieneDescripcion
      ) {
        feriadosHeaderRow =
          rowNumber

        feriadosHeaderMap =
          posiblesHeaders
      }
    }
  )
  if (
    feriadosHeaderRow !== null
  ) {
    const fechaColumn =
      feriadosHeaderMap.get(
        'FECHA'
      )
    const descripcionColumn =
      feriadosHeaderMap.get(
        'DESCRIPCION'
      ) ??
      feriadosHeaderMap.get(
        'FERIADO'
      )
    worksheetDias.eachRow(
      (row, rowNumber) => {
        if (
          rowNumber <=
          feriadosHeaderRow!
        ) {
          return
        }
        if (!fechaColumn) {
          return
        }
        const fecha =
          row.getCell(
            fechaColumn
          ).value
        if (!fecha) {
          return
        }
        const descripcion =
          descripcionColumn
            ? String(
                row.getCell(
                  descripcionColumn
                ).value ?? ''
              ).trim()
            : ''
        feriadosImportados.push({
          fecha,
          descripcion,
        })
      }
    )

    console.log(
      'FERIADOS EXTRAÍDOS:',
      feriadosImportados.length
    )

    console.log(
      'FERIADOS:',
      feriadosImportados
    )
  } else {
    console.warn(
      'No se encontró automáticamente una tabla de feriados en "DIAS".'
    )
  }

  console.log(
    '======================================'
  )
  console.log(
    'EXTRACCIÓN DE DIAS TERMINADA'
  )
  console.log(
    'Días mensuales:',
    daysByDni.size
  )
  console.log(
    'Trabajadores con incidencias:',
    incidenciasPorDni.size
  )
  console.log(
    'Feriados:',
    feriadosImportados.length
  )
  console.log(
    '======================================'
  )
}
  // ============================================================
  // PASO 8: GUARDAR LOS DÍAS EXTRAÍDOS
  // ============================================================
  setDiasImportados(
  daysByDni.size
)

setTrabajadores((current) => {
  const actualizados = current.map((trabajador) => {
    const clave = String(trabajador.dni)
      .trim()
      .padStart(8, '0')

    const diasTrabajados = daysByDni.get(clave)
    const incidencias =
      incidenciasPorDni.get(clave) ?? []

    const infoNombres =
      nombresPorDni.get(clave)

    return {
      ...trabajador,

      apellidoPaterno:
        trabajador.apellidoPaterno ||
        infoNombres?.apellidoPaterno ||
        '',

      apellidoMaterno:
        trabajador.apellidoMaterno ||
        infoNombres?.apellidoMaterno ||
        '',

      nombres:
        trabajador.nombres ||
        infoNombres?.nombres ||
        '',

      ...(diasTrabajados
        ? { diasTrabajados }
        : {}),

      incidencias,
    }
  })

  const nuevos = [...daysByDni.keys()]
    .filter(
      (dni) =>
        !current.some(
          (trabajador) =>
            String(trabajador.dni)
              .trim()
              .padStart(8, '0') === dni
        )
    )
    .map((dni) => {
      const infoNombres =
        nombresPorDni.get(dni)

      const diasTrabajados =
        daysByDni.get(dni)!

      const incidencias =
        incidenciasPorDni.get(dni) ?? []

      const mesesDias = months.filter(
        (mes) => diasTrabajados[mes] > 0
      )

      return {
        dni,
        apellidoPaterno:
          infoNombres?.apellidoPaterno ?? '',
        apellidoMaterno:
          infoNombres?.apellidoMaterno ?? '',
        nombres:
          infoNombres?.nombres ?? '',
        fechaInicio: null,
        fechaCese: null,
        remuneraciones: MESES_EN_CERO,
        diasTrabajados,
        incidencias,
        mesesDias,
        mesesPresentes: mesesDias,
      }
    })

  return [...actualizados, ...nuevos]
})

  setImportedFile(
    file.name
  )

  // ============================================================
  // PASO 10: RESULTADO DE ESTA PRIMERA ETAPA
  // ============================================================

  return
}
const worksheetRem = workbook.getWorksheet('REM')
        const worksheetIngresos = workbook.getWorksheet('INGRESOS 2025')
        const worksheet = worksheetRem || worksheetIngresos
        const esFormatoIngresos = !worksheetRem

        if (!worksheet) {
          alert('No se encontró la hoja "REM" ni "INGRESOS 2025" en el archivo.')
          return
        }

        const trabajadoresImportados: ImportedWorker[] = []

  const getExcelNumber = (value: ExcelJS.CellValue): number => {
    if (typeof value === 'number') {
      return value
    }

    if (
      typeof value === 'object' &&
      value !== null &&
      'result' in value &&
      typeof value.result === 'number'
    ) {
      return value.result
    }

    const numero = Number(value)

    return Number.isFinite(numero) ? numero : 0
  }

      const normalizeTituloRem = (valor: ExcelJS.CellValue): string =>
        String(typeof valor === 'object' && valor !== null && 'text' in valor ? valor.text ?? '' : valor ?? '')
          .normalize('NFD')
          .replace(/[\u0300-\u036f]/g, '')
          .trim()
          .replace(/\s+/g, ' ')
          .toUpperCase()

      // Formato "REM cabecera": hoja REM con los encabezados (DNI, apellidos,
      // nombres, meses) en la fila 1 y los datos desde la fila 2.
      const esFormatoRemCabeceras = (() => {
        if (esFormatoIngresos || !worksheetRem) return false
        const fila = worksheetRem.getRow(1)
        let tieneDni = false
        let mesesEncontrados = 0
        fila.eachCell((cell) => {
          const nombre = normalizeTituloRem(cell.value)
          if (nombre === 'DNI') tieneDni = true
          if (months.some((m) => m.toUpperCase() === nombre)) mesesEncontrados++
        })
        return tieneDni && mesesEncontrados >= 1
      })()

      const columnaMesRemCabeceras = new Map<string, number>()
      if (esFormatoRemCabeceras && worksheetRem) {
        worksheetRem.getRow(1).eachCell((cell, col) => {
          const nombre = normalizeTituloRem(cell.value)
          const mes = months.find((m) => m.toUpperCase() === nombre)
          if (mes) columnaMesRemCabeceras.set(mes, col)
        })
      }

      const buscarColumnaRemCabeceras = (nombres: string[], defecto: number): number => {
        if (!esFormatoRemCabeceras || !worksheetRem) return defecto
        let encontrada: number | null = null
        worksheetRem.getRow(1).eachCell((cell, col) => {
          if (encontrada === null && nombres.includes(normalizeTituloRem(cell.value))) encontrada = col
        })
        return encontrada ?? defecto
      }

      const colDniRemCabeceras = buscarColumnaRemCabeceras(['DNI'], 1)
      const colApRemCabeceras = buscarColumnaRemCabeceras(['APELLIDO PATERNO'], 2)
      const colAmRemCabeceras = buscarColumnaRemCabeceras(['APELLIDO MATERNO'], 3)
      const colNombresRemCabeceras = buscarColumnaRemCabeceras(['NOMBRES'], 4)
      const colInicioRemCabeceras = buscarColumnaRemCabeceras(['FECHA DE INICIO', 'FECHA INGRESO'], 5)
      const colCeseRemCabeceras = buscarColumnaRemCabeceras(['FECHA DE CESE', 'FECHA CESE'], 6)


      worksheet.eachRow((row: ExcelJS.Row) => {
    const dni = row.getCell(esFormatoRemCabeceras ? colDniRemCabeceras : 2).value

    // Solo procesamos filas que realmente tengan un DNI válido
  const dniTexto = String(dni ?? '')
    .trim()
    .replace(/\.0$/, '')

  if (!/^\d{7,9}$/.test(dniTexto)) {
    return
  }

  const dniFinal = /^\d{9}$/.test(dniTexto)
    ? dniTexto
    : dniTexto.padStart(8, '0')

 if (esFormatoIngresos) {
    const nombreCompleto = limpiarNombre(row.getCell(3).value)
    const partes = nombreCompleto.split(/\s+/).filter(Boolean)

    const remuneraciones: MonthlyValues = {
      enero: getExcelNumber(row.getCell(9).value),
      febrero: getExcelNumber(row.getCell(10).value),
      marzo: getExcelNumber(row.getCell(11).value),
      abril: getExcelNumber(row.getCell(12).value),
      mayo: getExcelNumber(row.getCell(13).value),
      junio: getExcelNumber(row.getCell(14).value),
      julio: getExcelNumber(row.getCell(15).value),
      agosto: getExcelNumber(row.getCell(16).value),
      septiembre: getExcelNumber(row.getCell(17).value),
      octubre: getExcelNumber(row.getCell(18).value),
      noviembre: getExcelNumber(row.getCell(19).value),
      diciembre: getExcelNumber(row.getCell(20).value),
      total: getExcelNumber(row.getCell(22).value),
    }

    const mesesRemuneraciones = months.filter(
      (mes) => remuneraciones[mes] > 0
    )

    trabajadoresImportados.push({
      dni: dniFinal,
      apellidoPaterno: partes[0] ?? '',
      apellidoMaterno: partes[1] ?? '',
      nombres: partes.slice(2).join(' ') || nombreCompleto,
      fechaInicio: row.getCell(4).value,
      fechaCese: row.getCell(5).value,

      remuneraciones,

      mesesRemuneraciones,

      mesesPresentes: mesesRemuneraciones,
    })
    return
  }

  if (esFormatoRemCabeceras) {
    const apellidoPaterno = limpiarNombre(row.getCell(colApRemCabeceras).value)
    const apellidoMaterno = limpiarNombre(row.getCell(colAmRemCabeceras).value)
    const nombres = limpiarNombre(row.getCell(colNombresRemCabeceras).value)
    const fechaInicio = row.getCell(colInicioRemCabeceras).value
    const fechaCese = row.getCell(colCeseRemCabeceras).value

    const remuneraciones: MonthlyValues = { ...MESES_EN_CERO }
    for (const [mes, col] of columnaMesRemCabeceras) {
      remuneraciones[mes as keyof MonthlyValues] = getExcelNumber(row.getCell(col).value)
    }
    remuneraciones.total = months.reduce((sum, mes) => sum + remuneraciones[mes], 0)

    const mesesRemuneraciones = months.filter(
      (mes) => remuneraciones[mes] > 0
    )

    trabajadoresImportados.push({
      dni: dniFinal,
      apellidoPaterno,
      apellidoMaterno,
      nombres,
      fechaInicio,
      fechaCese,
      remuneraciones,
      mesesRemuneraciones,
      mesesPresentes: mesesRemuneraciones,
    })
    return
  }

 const apellidoPaterno = limpiarNombre(row.getCell(3).value)
const apellidoMaterno = limpiarNombre(row.getCell(4).value)
const nombres = limpiarNombre(row.getCell(5).value)

const fechaInicio = row.getCell(6).value
const fechaCese = row.getCell(7).value

    const remuneraciones: MonthlyValues = {
  enero: getExcelNumber(row.getCell(9).value),
  febrero: getExcelNumber(row.getCell(10).value),
  marzo: getExcelNumber(row.getCell(11).value),
  abril: getExcelNumber(row.getCell(12).value),
  mayo: getExcelNumber(row.getCell(13).value),
  junio: getExcelNumber(row.getCell(14).value),
  julio: getExcelNumber(row.getCell(15).value),
  agosto: getExcelNumber(row.getCell(16).value),
  septiembre: getExcelNumber(row.getCell(17).value),
  octubre: getExcelNumber(row.getCell(18).value),
  noviembre: getExcelNumber(row.getCell(19).value),
  diciembre: getExcelNumber(row.getCell(20).value),
  total: getExcelNumber(row.getCell(21).value),
}

// Detectar los meses que realmente tienen remuneración registrada
const mesesRemuneraciones = months.filter(
  (mes) => remuneraciones[mes] > 0
)

trabajadoresImportados.push({
  dni: dniFinal,
  apellidoPaterno,
  apellidoMaterno,
  nombres,
  fechaInicio,
  fechaCese,

  remuneraciones,

  mesesRemuneraciones,

  // En esta etapa solo existe información REM.
  // DÍAS se agregará cuando se importe el Excel de días.
  mesesPresentes: mesesRemuneraciones,
})
  })
  console.log(
  '[TRAZABILIDAD REM] Trabajador importado antes de consolidar:',
  trabajadoresImportados[0]?.dni,
  {
    remuneraciones: trabajadoresImportados[0]?.remuneraciones,
    mesesRemuneraciones: trabajadoresImportados[0]?.mesesRemuneraciones,
    mesesPresentes: trabajadoresImportados[0]?.mesesPresentes,
  }
)
  setTrabajadores((current) => {
  return trabajadoresImportados.map((rem) => {
    const existente = current.find(
      (trabajador) =>
        String(trabajador.dni).trim().padStart(8, '0') === rem.dni
    )

    if (!existente) {
      return rem
    }

    return {
      ...existente,
      ...rem,
      diasTrabajados: existente.diasTrabajados,
      mesesDias: existente.mesesDias,
      mesesPresentes: Array.from(
        new Set([
          ...(rem.mesesRemuneraciones ?? []),
          ...(existente.mesesDias ?? []),
        ])
      ),
    }
  })
})
  setImportedFile(file.name)
  alert(
    `Excel leído correctamente.\n\nTrabajadores encontrados: ${trabajadoresImportados.length}`
  )
        } catch (error) {
        console.error('Error al leer Excel:', error)
        alert(`No se pudo leer el archivo Excel.${error instanceof Error ? `\n\n${error.message}` : ''}`)
      } finally {
        setLoading(false)
      }
    }
    const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0]
      if (file) {
        handleFile(file)
      }
      event.target.value = ''
    }

    const trabajadoresFiltrados = useMemo(() => {
      const texto = busqueda.toLowerCase().trim()
      return trabajadores
        .map((t, idx) => ({ trabajador: t, idx }))
        .filter(({ trabajador: t }) => {
          const coincideBusqueda = !texto ||
            t.dni.includes(texto) ||
            t.apellidoPaterno.toLowerCase().includes(texto) ||
            t.apellidoMaterno.toLowerCase().includes(texto) ||
            t.nombres.toLowerCase().includes(texto) ||
            `${t.apellidoPaterno} ${t.apellidoMaterno} ${t.nombres}`.toLowerCase().includes(texto)
          let coincideFiltro = true
          if (filtroEstado === 'conCese') {
            coincideFiltro = !!t.fechaCese && String(t.fechaCese).trim() !== ''
          } else if (filtroEstado === 'activos') {
            coincideFiltro = !t.fechaCese || String(t.fechaCese).trim() === ''
          }
          return coincideBusqueda && coincideFiltro
        })
    }, [trabajadores, busqueda, filtroEstado])

    return (
      <Card className="shadow-none">
        <CardHeader className="flex flex-row items-start justify-between gap-4">
          <div>
            <CardTitle className="text-base">{title}</CardTitle>
            <CardDescription>{description}</CardDescription>
          </div>
          <Badge variant="outline">
            XLSX / CSV
          </Badge>
        </CardHeader>
        <CardContent>
    {!importedFile ? (
      <>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={handleInputChange}
        />
        <button
          type="button"
          disabled={loading}
          onClick={() => inputRef.current?.click()}
          className="group flex w-full flex-col items-center gap-4 rounded-2xl border-2 border-dashed border-primary/30 bg-gradient-to-b from-primary/5 to-transparent px-6 py-10 text-center transition-colors hover:border-primary/50 hover:from-primary/10 disabled:cursor-not-allowed disabled:opacity-60"
        >
          <div className="flex size-14 items-center justify-center rounded-2xl bg-primary/10 text-primary shadow-sm transition-transform group-hover:scale-105">
            <FileSpreadsheet className="size-7" />
          </div>
          {loading ? (
            <span className="text-base font-semibold">Leyendo archivo...</span>
          ) : (
            <>
              <span className="text-base font-semibold">
                Importar {title}
              </span>
              <span className="max-w-sm text-sm text-muted-foreground">
                {isRemunerations
                  ? 'Carga la información mensual o consolidada de remuneraciones.'
                  : 'Carga los días trabajados, ausencias y licencias del ejercicio.'}
              </span>
            </>
          )}
          <span className="mt-1 inline-flex items-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-semibold text-primary-foreground shadow-sm transition-shadow group-hover:shadow-md">
            <Upload className="size-4" />
            {loading ? 'Procesando...' : 'Importar'}
          </span>
        </button>
      </>
    ) : (
      <div className="space-y-4">
        {/* Archivo cargado */}
        <div className="flex items-center justify-between gap-3 rounded-2xl border border-emerald-500/30 bg-emerald-500/5 p-4">
          <div className="flex items-center gap-3">
            <div className="flex size-10 items-center justify-center rounded-xl bg-emerald-500/15">
              <Check className="size-5 text-emerald-400" />
            </div>
            <div>
              <p className="font-medium">
                {importedFile}
              </p>
              <p className="text-sm text-emerald-500">
                {isRemunerations ? trabajadores.length : diasImportados} trabajadores detectados · Mapeo automático disponible
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={() => inputRef.current?.click()}
            className="rounded-lg border bg-background px-3 py-2 text-sm font-medium transition-colors hover:bg-muted"
          >
            Cambiar archivo
          </button>
        </div>
        {/* Resumen */}
        <div className="grid gap-3 md:grid-cols-3">
         <div className="rounded-2xl border p-4">
           <p className="text-sm text-muted-foreground">
             Trabajadores detectados
           </p>
           <p className="mt-2 text-2xl font-semibold">
             {isRemunerations ? trabajadores.length : diasImportados}
           </p>
         </div>
         <div className="rounded-2xl border p-4">
  <p className="text-sm text-muted-foreground">
    {isRemunerations ? 'Remuneración anual' : 'Días registrados'}
  </p>
  <p className="mt-2 text-2xl font-semibold">
    {isRemunerations && trabajadores.length > 0
      ? `S/ ${trabajadores
          .reduce((total, trabajador) => {
            return total + trabajador.remuneraciones.total
          }, 0)
          .toLocaleString('es-PE', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}`
      : !isRemunerations && trabajadores.length > 0
        ? `${trabajadores.reduce((total, trabajador) => {
            return total + (trabajador.diasTrabajados?.total ?? 0)
          }, 0)} días`
        : '—'}
  </p>
</div>
          <div className="rounded-2xl border p-4">
            <p className="text-sm text-muted-foreground">
              Observaciones
            </p>
            <p className="mt-2 text-2xl font-semibold">
              0
            </p>
          </div>
        </div>
  {/* Buscador y filtros */}
  <div className="flex flex-col gap-3 sm:flex-row sm:items-center">
    <div className="relative flex-1">
      <Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />
      <Input
        value={busqueda}
        onChange={(e) => setBusqueda(e.target.value)}
        placeholder="Buscar por nombre o DNI..."
        className="pl-9"
      />
      {busqueda && (
        <button
          type="button"
          onClick={() => setBusqueda('')}
          className="absolute right-3 top-1/2 -translate-y-1/2 text-muted-foreground hover:text-foreground"
        >
          <X className="size-4" />
        </button>
      )}
    </div>
    <Select value={filtroEstado} onValueChange={(value) => setFiltroEstado(value ?? 'todos')}>
      <SelectTrigger className="w-full sm:w-48">
        <Filter className="mr-2 size-4" />
        <SelectValue placeholder="Filtrar por estado" />
      </SelectTrigger>
      <SelectContent>
        <SelectItem value="todos">Todos</SelectItem>
        <SelectItem value="activos">Activos (sin cese)</SelectItem>
        <SelectItem value="conCese">Con fecha de cese</SelectItem>

      </SelectContent>
    </Select>
  </div>
  {busqueda || filtroEstado !== 'todos' ? (
    <p className="text-xs text-muted-foreground">
      Mostrando {trabajadoresFiltrados.length} de {trabajadores.length} trabajadores
    </p>
  ) : null}
  {/* Trabajadores */}
  <div className="overflow-hidden rounded-2xl border">
<div className={`grid gap-4 border-b bg-muted/20 px-4 py-3 text-sm font-medium ${isRemunerations ? 'grid-cols-4' : 'grid-cols-[minmax(0,1.5fr)_minmax(5.5rem,0.7fr)_minmax(6rem,0.8fr)_minmax(6rem,0.8fr)]'}`}>
      <span>Trabajador</span>
      <span>DNI</span>
      <span>Inicio</span>
      <span>Cese</span>
    </div>
    <div className="divide-y">
      {trabajadoresFiltrados.length === 0 && (
        <div className="px-4 py-10 text-center text-sm text-muted-foreground">
          No se encontraron trabajadores con los filtros aplicados.
        </div>
      )}
      {trabajadoresFiltrados.map(({ trabajador, idx }) => (
        <div key={trabajador.dni}>
          {/* FILA DEL TRABAJADOR */}
          <div
            onClick={() =>
              setTrabajadorSeleccionado(
                trabajadorSeleccionado === idx ? null : idx
              )
            }
            className={`grid gap-4 px-4 py-3 text-sm cursor-pointer hover:bg-muted/30 ${isRemunerations ? 'grid-cols-4' : 'grid-cols-[minmax(0,1.5fr)_minmax(5.5rem,0.7fr)_minmax(6rem,0.8fr)_minmax(6rem,0.8fr)]'}`}
          >
            <div className="min-w-0">
              <p className="font-medium">
                {trabajador.apellidoPaterno}{" "}
                {trabajador.apellidoMaterno}
              </p>
              <p className="text-xs text-muted-foreground">
                {trabajador.nombres}
              </p>
            </div>
            <span className="min-w-0 overflow-hidden text-ellipsis whitespace-nowrap">
              {trabajador.dni}
            </span>
            <span className="min-w-0 whitespace-nowrap">
              {fechaIngresoTrabajador(trabajador).etiqueta}
            </span>
            <span className="min-w-0 whitespace-nowrap">
              {fechaCeseTrabajador(trabajador).etiqueta}
            </span>
          </div>
          {/* DETALLE DEL TRABAJADOR */}
          {trabajadorSeleccionado === idx && (
            <div className="border-t bg-muted/10 px-4 py-4">
              {!isRemunerations && <PeriodosTrabajados worker={trabajador} />}

              {isRemunerations && (
  <>
    <p className="mb-3 text-sm font-semibold">
      Remuneraciones mensuales
    </p>

                <div className="grid grid-cols-2 gap-3 md:grid-cols-4">

                  {[
                    ["Enero", trabajador.remuneraciones.enero],
                    ["Febrero", trabajador.remuneraciones.febrero],
                    ["Marzo", trabajador.remuneraciones.marzo],
                    ["Abril", trabajador.remuneraciones.abril],
                    ["Mayo", trabajador.remuneraciones.mayo],
                    ["Junio", trabajador.remuneraciones.junio],
                    ["Julio", trabajador.remuneraciones.julio],
                    ["Agosto", trabajador.remuneraciones.agosto],
                    ["Septiembre", trabajador.remuneraciones.septiembre],
                    ["Octubre", trabajador.remuneraciones.octubre],
                    ["Noviembre", trabajador.remuneraciones.noviembre],
                    ["Diciembre", trabajador.remuneraciones.diciembre],
                  ].map(([mes, monto]) => (
                    <div
                      key={mes}
                      className="rounded-lg border border-border/60 bg-background p-3"
                    >
                      <p className="text-xs text-muted-foreground">
                        {mes}
                      </p>
                      <p className="mt-1 text-sm font-semibold">
                        S/{" "}
                        {Number(monto).toLocaleString("es-PE", {
                          minimumFractionDigits: 2,
                          maximumFractionDigits: 2,
                        })}
                      </p>
                    </div>
                  ))}
                </div>
<div className="mt-3 flex items-center justify-between rounded-lg border border-primary/20 bg-primary/5 p-3"> 
                <span className="text-sm font-medium">
                  Remuneración total anual
                </span>
                <span className="font-semibold text-primary">
                  S/{" "}
                  {Number(trabajador.remuneraciones.total).toLocaleString(
                    "es-PE",
                    {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    }
                  )}
                </span>
              </div>
    </>
  )}
              {!isRemunerations && trabajador.diasTrabajados && (
  <div className="mt-5 rounded-xl border border-border/60 bg-muted/10 p-4">
    <div className="mb-4 flex items-center justify-between">
      <div>
        <p className="text-sm font-semibold">
          Días trabajados
        </p>
        <p className="text-xs text-muted-foreground">
          Resumen mensual registrado
        </p>
      </div>
      <div className="flex items-center gap-2">
        {getPeriodosTrabajados(trabajador).some((p) => p.estado === 'Reingresó') && (
          <Badge className="border-amber-500/30 bg-amber-500/10 text-amber-500">
            Reingreso detectado
          </Badge>
        )}
        <div className="rounded-lg border border-primary/20 bg-primary/10 px-3 py-1.5">
          <span className="text-sm font-semibold text-primary">
            {trabajador.diasTrabajados.total} días
          </span>
        </div>
      </div>
    </div>
    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 xl:grid-cols-6">
      {[
        ['Enero', trabajador.diasTrabajados.enero],
        ['Febrero', trabajador.diasTrabajados.febrero],
        ['Marzo', trabajador.diasTrabajados.marzo],
        ['Abril', trabajador.diasTrabajados.abril],
        ['Mayo', trabajador.diasTrabajados.mayo],
        ['Junio', trabajador.diasTrabajados.junio],
        ['Julio', trabajador.diasTrabajados.julio],
        ['Agosto', trabajador.diasTrabajados.agosto],
        ['Septiembre', trabajador.diasTrabajados.septiembre],
        ['Octubre', trabajador.diasTrabajados.octubre],
        ['Noviembre', trabajador.diasTrabajados.noviembre],
        ['Diciembre', trabajador.diasTrabajados.diciembre],
      ].map(([mes, dias]) => (
        <div
          key={mes}
          className={`min-w-0 rounded-lg border border-border/60 bg-background p-3 text-center transition-colors hover:border-primary/30 hover:bg-primary/5 ${
            dias === 0 ? 'opacity-50' : ''
          }`}
        >
          <p className="text-[11px] font-medium text-muted-foreground">
            {mes}
          </p>
          <p className="mt-1 text-lg font-semibold">
            {dias}
          </p>
          <p className="text-[10px] text-muted-foreground">
            días
          </p>
        </div>
      ))}
    </div>
<div className="mt-3 flex items-center justify-between rounded-lg border border-border/60 bg-background px-4 py-3">
      <span className="text-xs font-medium text-muted-foreground">
        Total anual registrado
      </span>
      <span className="text-sm font-bold text-primary">
        {trabajador.diasTrabajados.total} días
      </span>
    </div>
    <div className="mt-4">
      <Button
        variant="outline"
        size="sm"
        onClick={() =>
          setIncidenciaTrabajador(
            incidenciaTrabajador === idx ? null : idx
          )
        }
      >
        Incidencia
      </Button>
    </div>
    {incidenciaTrabajador === idx && (
      <div className="mt-3 border-t border-border/60 pt-3">
        <p className="mb-2 text-sm font-semibold">Incidencias</p>
        <p className="mb-3 text-xs text-muted-foreground">
          
        </p>

        {trabajador.incidencias && trabajador.incidencias.length > 0 ? (
          <div className="space-y-2">
            {trabajador.incidencias.map((incidencia, incidenciaIndex) => {
              const periodo = formatearPeriodo(incidencia.periodo)
              const rango = `${formatExcelDate(incidencia.fechaInicio)} al ${formatExcelDate(incidencia.fechaFin)}`
              const dias = `${incidencia.cantidadDias} ${incidencia.cantidadDias === 1 ? 'día' : 'días'}`

              return (
                <div
                  key={`${incidencia.codigo}-${incidenciaIndex}`}
                  className="rounded-lg border border-border/60 bg-background p-3 text-sm"
                >
                  <p className="font-medium">
                    {incidencia.descripcion || incidencia.codigo || 'Incidencia registrada'}
                  </p>
                  <p className="mt-1 text-xs text-muted-foreground">
                    {periodo ? `${periodo} · ${dias}` : `${rango} · ${dias}`}
                  </p>
                  {!periodo && !formatExcelDate(incidencia.fechaInicio) && (
                    <p className="mt-1 text-[11px] text-amber-400">
                      Período no disponible en el archivo: queda pendiente de identificar.
                    </p>
                  )}
                </div>
              )
            })}
          </div>
        ) : (
          <p className="text-sm text-muted-foreground">
            No tiene incidencias registradas.
          </p>
        )}
      </div>
    )}
          </div>
      )}
            </div>
          )}
        </div>
        ))}
    </div>
  </div>
      </div>
    )}
  </CardContent>
      </Card>
    )
  }
  export function CalculationView({ results, totals, errors, onRun, onSelect, onSummaryChange, onReprocessed }: { results: EmployeeUtilityResult[]; totals: ReturnType<typeof getTotals>; errors: string[]; onRun: () => void; onSelect: (result: EmployeeUtilityResult) => void; onSummaryChange?: (summary: ImportSummary) => void; onReprocessed?: (employees: Employee[]) => void }) { const { parameters } = useSettings(); const top = [...results].sort((a,b) => b.finalAmount - a.finalAmount).slice(0,5);
  const totalDiasEfectivos = results.reduce((suma, r) => suma + (r.diasEfectivos ?? 0), 0)
  const totalRemuneraciones = results.reduce((suma, r) => suma + (r.remuneracionComputable ?? 0), 0)
  const fondoPorDias = totals.fund / 2
  const fondoPorRemuneraciones = totals.fund / 2
  const sumaImporteDias = results.reduce((suma, r) => suma + r.distribucion.utilidadPorDias, 0)
  const sumaImporteRemuneraciones = results.reduce((suma, r) => suma + r.distribucion.utilidadPorRemuneraciones, 0)
  const sumaUtilidadBruta = results.reduce((suma, r) => suma + r.distribucion.utilidadBruta, 0)
  const diffDias = Math.round((sumaImporteDias - fondoPorDias) * 100) / 100
  const diffRemuneraciones = Math.round((sumaImporteRemuneraciones - fondoPorRemuneraciones) * 100) / 100
  const diffBruta = Math.round((sumaUtilidadBruta - totals.fund) * 100) / 100
  const diasEfectivosNegativos = results.filter((r) => (r.diasEfectivos ?? 0) < 0).length
  const conDivisionCero = totalDiasEfectivos <= 0 || totalRemuneraciones <= 0 || totals.fund <= 0
  const [trabajadores, setTrabajadores] = useState<ImportedWorker[]>([])
  const [trabajadoresDias, setTrabajadoresDias] = useState<ImportedWorker[]>([])
  const [jornadas, setJornadas] = useState<Record<string, Jornada>>({})
  const [fechasEditadas, setFechasEditadas] = useState<Record<string, { fechaInicio?: string; fechaCese?: string }>>({})
  const [mesesEditados, setMesesEditados] = useState<Record<string, string[]>>({})
  const [incidenciasEditadas, setIncidenciasEditadas] = useState<Record<string, IncidenciaImportada[]>>({})
  const handleReprocesar = () => {
    const empleados: Employee[] = trabajadoresUnificados.map((t) => {
      const jornada = (jornadas[t.dni] ?? t.jornada) as Jornada | undefined
      const incidencias = incidenciasEditadas[t.dni] ?? t.incidencias ?? []
      const noLaborados = calcularDiasNoLaborados(
        incidenciasAplicables(incidencias).map(inc => ({
          cantidadDias: inc.cantidadDias,
          diasNeto: inc.diasNeto,
          fechaInicio: parseIncFecha(inc.fechaInicio),
          fechaFin: parseIncFecha(inc.fechaFin),
        })),
        jornada ?? 5,
        new Set(t.feriados ?? [])
      )
      const mesesConDias = new Set(mesesTrabajados(t, mesesEditados[t.dni]))
      // DÍAS POSIBLES / LABORABLES: prioridad a los días posibles del formato
      // CITIKOLD (columna L); en otros formatos se usan los días por mes cargados.
      const posibles = t.diasPosibles
        ?? (t.laborablesPorMes
          ? months.reduce((sum, mes) => sum + (mesesConDias.has(mes) ? (t.laborablesPorMes?.[mes] ?? 0) : 0), 0)
          : (t.diasTrabajados?.total ?? 0))
      // DÍAS EFECTIVOS = DÍAS POSIBLES − DÍAS NO LABORADOS (un solo descuento)
      const efectivos = Math.max(0, posibles - noLaborados)
      const remuneracionComputable = t.remuneraciones.total ?? 0
      return {
        id: t.dni,
        code: t.dni,
        name: `${t.apellidoPaterno} ${t.apellidoMaterno}`.trim() || t.nombres,
        role: '',
        department: '',
        days: efectivos,
        diasLaborables: posibles,
        diasNoLaborados: noLaborados,
        diasEfectivos: efectivos,
        jornada,
        remuneration: remuneracionComputable,
        remuneracionComputable,
        status: posibles > 0 && remuneracionComputable > 0 ? 'Completo' : 'Pendiente',
      } as Employee & { remuneracionComputable: number }
    })
    setReprocesado(true)
    onReprocessed?.(empleados)
  }
  const [activeStep, setActiveStep] = useState(0)
  const [detalleDni, setDetalleDni] = useState<string | null>(null)
  const [reprocesado, setReprocesado] = useState(false)
  const listoParaProcesar = trabajadores.length > 0 && trabajadoresDias.length > 0
  // Consolidado de la carga de días y la de remuneraciones: la misma función
  // que usa el motor, para que Validación no tenga una segunda fuente de datos.
  const trabajadoresUnificados = useMemo(() => consolidarTrabajadores(trabajadoresDias, trabajadores), [trabajadores, trabajadoresDias])

  const resultadosCoinciden =
    results.length > 0 &&
    results.length === trabajadoresUnificados.length &&
    trabajadoresUnificados.every(
      (t, i) => String(results[i]?.id).padStart(8, '0') === String(t.dni).padStart(8, '0')
    )
  const distribucionLista = reprocesado && resultadosCoinciden

  const handleProcesarAhora = () => {
    if (listoParaProcesar) {
      setActiveStep(1)
    }
  }

  useEffect(() => {
    onSummaryChange?.({
      remuneracionesTrabajadores: trabajadores.length,
      totalRemuneracionAnual: trabajadores.reduce((total, trabajador) => total + trabajador.remuneraciones.total, 0),
      diasTrabajadores: trabajadoresDias.length,
      totalDiasRegistrados: trabajadoresDias.reduce((total, trabajador) => total + (trabajador.diasTrabajados?.total ?? 0), 0),
      totalIncidencias: trabajadoresDias.reduce((total, trabajador) => total + (trabajador.incidencias?.length ?? 0), 0),
    })
  }, [trabajadores, trabajadoresDias, onSummaryChange])
  ; return (
  <div className="flex flex-col gap-7">
<SectionTitle
  eyebrow="Cálculo · Paso 2 de 3"
  title="Distribución de utilidades"
  description="Revisa los parámetros, valida los factores de distribución y genera el cálculo anual del ejercicio."
  action={
    <div className="flex gap-2 sm:hidden">
      <Button variant="outline" size="sm">
        <Check className="mr-2 size-4" />
        Guardar
      </Button>

      <Button size="sm" onClick={onRun}>
        <Play className="mr-2 size-4" />
        Ejecutar
      </Button>
    </div>
  }
/>
    {/* Flujo del cálculo */}
<div className="rounded-xl border border-border/60 bg-card p-5">
  <div className="mb-5 flex items-center justify-between">
    <div>
      <p className="text-sm font-semibold">Flujo del cálculo</p>
      <p className="mt-1 text-xs text-muted-foreground">
        Proceso de cálculo de utilidades
      </p>
    </div>
    <span className="text-xs font-medium text-primary">
      Paso {activeStep + 1} de 3
    </span>
  </div>
  <div className="flex items-center">
    {[
      "Carga",
      "Validación",
      "Distribución",
    ].map((label, index) => (
      <div key={label} className="flex flex-1 items-center">
        <button
          type="button"
          onClick={() => { if (index === 2 && !distribucionLista) return; setActiveStep(index) }}
          className={`flex min-w-0 flex-col items-center transition-opacity ${index === 2 && !distribucionLista ? 'cursor-not-allowed opacity-40' : 'hover:opacity-80'}`}
        >
          <div
            className={`flex size-9 items-center justify-center rounded-full border text-sm font-semibold transition-colors ${
              index === activeStep
                ? "border-primary bg-primary text-primary-foreground"
                : "border-border bg-muted/30 text-muted-foreground hover:border-primary/50 hover:text-primary"
            }`}
          >
            {index + 1}
          </div>

          <span
            className={`mt-2 text-center text-xs font-medium ${
              index === activeStep
                ? "text-primary"
                : "text-muted-foreground hover:text-primary"
            }`}
          >
            {label}
          </span>
        </button>
        {index < 2 && (
          <div className="mx-2 h-px flex-1 bg-border" />
        )}
      </div>
    ))}
  </div>
</div>
{activeStep === 1 ? (
  <div className="flex flex-col gap-6">
    <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
      <div>
        <p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">
          Validación
        </p>
        <h2 className="text-xl font-semibold tracking-tight">
          Validación de trabajadores
        </h2>
        <p className="mt-1 text-sm text-muted-foreground">
          Revisa la información consolidada de remuneraciones y días, corrige jornada, períodos, fechas e incidencias y reprocesa los días y resultados.
        </p>
      </div>
      <BotonExportarInformacion
        onClick={() => { void exportarValidacion(trabajadoresUnificados, jornadas, fechasEditadas, mesesEditados, incidenciasEditadas) }}
        disabled={trabajadoresUnificados.length === 0}
      />
    </div>

    <PanelValidacion
      trabajadores={trabajadoresUnificados}
      jornadas={jornadas}
      setJornadas={setJornadas}
      fechasEditadas={fechasEditadas}
      setFechasEditadas={setFechasEditadas}
      mesesEditados={mesesEditados}
      setMesesEditados={setMesesEditados}
      incidenciasEditadas={incidenciasEditadas}
      setIncidenciasEditadas={setIncidenciasEditadas}
      onReprocesar={handleReprocesar}
    />

    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
      <p className="text-xs text-muted-foreground">
        Los datos de remuneraciones y días llegan automáticamente desde el apartado de Carga.
      </p>
      <div className="flex items-center gap-2">
        <Button variant="outline" onClick={() => setActiveStep(0)}>Volver a carga</Button>
        <Button onClick={handleReprocesar} disabled={trabajadoresUnificados.length === 0}>
          <Play className="mr-2 size-4" />Reprocesar días y resultados
        </Button>
        <Button variant="outline" onClick={() => setActiveStep(2)} disabled={!distribucionLista}>
          Siguiente: Distribución<ArrowRight data-icon="inline-end" />
        </Button>
      </div>
    </div>
  </div>
) : activeStep === 2 ? (
  <div className="flex flex-col gap-6">
    <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
      <div>
        <p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">
          Distribución
        </p>
        <h2 className="text-xl font-semibold tracking-tight">
          Distribución del fondo
        </h2>
        <p className="mt-1 text-sm text-muted-foreground">
          Resumen del fondo legal, factores y resultados del ejercicio.
        </p>
      </div>
      <BotonExportarInformacion
        onClick={() => { void exportarDistribucion(results, totals) }}
        disabled={results.length === 0}
      />
    </div>

    <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
  <Kpi
    label="Fondo total"
    value={money.format(totals.fund)}
    hint="10% de la renta neta"
    icon={ShieldCheck}
    tone="teal"
  />
  <Kpi
    label="Fondo por días (50%)"
    value={money.format(fondoPorDias)}
    hint="según días efectivos"
    icon={CalendarDays}
  />
  <Kpi
    label="Fondo por remuneraciones (50%)"
    value={money.format(fondoPorRemuneraciones)}
    hint="según remuneración computable"
    icon={Calculator}
  />
  <Kpi
    label="Total días efectivos"
    value={number.format(totalDiasEfectivos)}
    hint="Σ días efectivos del ejercicio"
    icon={ClipboardCheck}
  />
  <Kpi
    label="Total remuneraciones"
    value={money.format(totalRemuneraciones)}
    hint="Σ remuneraciones computables"
    icon={ArrowDownToLine}
    tone="amber"
  />
  <Kpi
    label="Trabajadores"
    value={number.format(results.length)}
    hint={`${results.length} registros procesados`}
    icon={Users}
    tone="amber"
  />
  <Kpi
    label="A distribuir"
    value={money.format(totals.distributed)}
    hint="tras tope y quinta categoría"
    icon={ArrowDownToLine}
  />
  <Kpi
    label="Remanente"
    value={money.format(totals.remainder)}
    hint={`${totals.capped} topes aplicados`}
    icon={AlertCircle}
    tone="rose"
  />
  </div>

    <Card>
      <CardHeader>
        <CardTitle>Comprobaciones del reparto</CardTitle>
        <CardDescription>
          Verifica que los importes por días, por remuneración y la utilidad bruta sumada cuadren con los fondos calculados, usando las reglas de redondeo actuales.
        </CardDescription>
      </CardHeader>
      <CardContent>
        <div className="grid gap-3 sm:grid-cols-2">
          <div className={`flex items-start gap-3 rounded-lg border p-3 ${diffDias === 0 ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-amber-500/30 bg-amber-500/10'}`}>
            <div className={`mt-0.5 rounded-md p-1.5 ${diffDias === 0 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-amber-500/10 text-amber-500'}`}>
              {diffDias === 0 ? <Check className="size-4" /> : <AlertCircle className="size-4" />}
            </div>
            <div>
              <p className="text-sm font-medium">Suma importe por días</p>
              <p className="mt-0.5 text-xs text-muted-foreground">
                {money.format(sumaImporteDias)} vs fondo por días {money.format(fondoPorDias)}
              </p>
              <p className={`mt-0.5 text-xs ${diffDias === 0 ? 'text-emerald-500' : 'text-amber-500'}`}>
                {diffDias === 0 ? 'Cuadra exacto' : `Diferencia por redondeo: ${money.format(diffDias)}`}
              </p>
            </div>
          </div>
          <div className={`flex items-start gap-3 rounded-lg border p-3 ${diffRemuneraciones === 0 ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-amber-500/30 bg-amber-500/10'}`}>
            <div className={`mt-0.5 rounded-md p-1.5 ${diffRemuneraciones === 0 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-amber-500/10 text-amber-500'}`}>
              {diffRemuneraciones === 0 ? <Check className="size-4" /> : <AlertCircle className="size-4" />}
            </div>
            <div>
              <p className="text-sm font-medium">Suma importe por remuneraciones</p>
              <p className="mt-0.5 text-xs text-muted-foreground">
                {money.format(sumaImporteRemuneraciones)} vs fondo por remuneraciones {money.format(fondoPorRemuneraciones)}
              </p>
              <p className={`mt-0.5 text-xs ${diffRemuneraciones === 0 ? 'text-emerald-500' : 'text-amber-500'}`}>
                {diffRemuneraciones === 0 ? 'Cuadra exacto' : `Diferencia por redondeo: ${money.format(diffRemuneraciones)}`}
              </p>
            </div>
          </div>
          <div className={`flex items-start gap-3 rounded-lg border p-3 ${diffBruta === 0 ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-amber-500/30 bg-amber-500/10'}`}>
            <div className={`mt-0.5 rounded-md p-1.5 ${diffBruta === 0 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-amber-500/10 text-amber-500'}`}>
              {diffBruta === 0 ? <Check className="size-4" /> : <AlertCircle className="size-4" />}
            </div>
            <div>
              <p className="text-sm font-medium">Suma utilidad bruta vs fondo</p>
              <p className="mt-0.5 text-xs text-muted-foreground">
                {money.format(sumaUtilidadBruta)} vs fondo total {money.format(totals.fund)}
              </p>
              <p className={`mt-0.5 text-xs ${diffBruta === 0 ? 'text-emerald-500' : 'text-amber-500'}`}>
                {diffBruta === 0 ? 'Cuadra exacto (antes de tope y quinta)' : `Diferencia por redondeo: ${money.format(diffBruta)}`}
              </p>
            </div>
          </div>
          <div className={`flex items-start gap-3 rounded-lg border p-3 ${diasEfectivosNegativos === 0 ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-rose-500/30 bg-rose-500/10'}`}>
            <div className={`mt-0.5 rounded-md p-1.5 ${diasEfectivosNegativos === 0 ? 'bg-emerald-500/10 text-emerald-500' : 'bg-rose-500/10 text-rose-500'}`}>
              {diasEfectivosNegativos === 0 ? <Check className="size-4" /> : <AlertCircle className="size-4" />}
            </div>
            <div>
              <p className="text-sm font-medium">Días efectivos negativos</p>
              <p className={`mt-0.5 text-xs ${diasEfectivosNegativos === 0 ? 'text-emerald-500' : 'text-rose-500'}`}>
                {diasEfectivosNegativos === 0 ? 'Ningún trabajador con días efectivos negativos' : `${diasEfectivosNegativos} trabajador(es) con días efectivos negativos`}
              </p>
            </div>
          </div>
          <div className={`flex items-start gap-3 rounded-lg border p-3 ${!conDivisionCero ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-rose-500/30 bg-rose-500/10'}`}>
            <div className={`mt-0.5 rounded-md p-1.5 ${!conDivisionCero ? 'bg-emerald-500/10 text-emerald-500' : 'bg-rose-500/10 text-rose-500'}`}>
              {!conDivisionCero ? <Check className="size-4" /> : <AlertCircle className="size-4" />}
            </div>
            <div>
              <p className="text-sm font-medium">Divisiones entre cero</p>
              <p className={`mt-0.5 text-xs ${!conDivisionCero ? 'text-emerald-500' : 'text-rose-500'}`}>
                {!conDivisionCero ? 'Sin divisiones entre cero en los factores' : 'Algún divisor es cero (total días, remuneraciones o fondo)'}
              </p>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>

    {errors.length > 0 && (
    <Alert variant="destructive">
      <AlertCircle className="size-4" />
      <AlertTitle>
        Validaciones pendientes
      </AlertTitle>
      <AlertDescription>
        {errors.join(' ')}
      </AlertDescription>
    </Alert>
  )}

    <div className="grid gap-5 xl:grid-cols-[1.15fr_.85fr]">
    {/* RESUMEN DEL FONDO */}
    <Card>
      <CardHeader>
        <div className="flex items-start justify-between gap-4">
          <div>
            <CardTitle>
              Resumen del fondo
            </CardTitle>
            <CardDescription>
              Distribución legal aplicable al ejercicio 2026
            </CardDescription>
          </div>
          {errors.length === 0 && (
            <Badge
              variant="secondary"
              className="bg-emerald-500/10 text-emerald-400"
            >
              Validado
            </Badge>
          )}
        </div>
      </CardHeader>
      <CardContent>
        <div className="flex flex-col gap-5">
          <div className="grid gap-4 sm:grid-cols-3">
            <div>
              <p className="text-xs text-muted-foreground">
                Renta neta
              </p>
              <p className="mt-1 text-lg font-semibold">
                {money.format(parameters.income)}
              </p>
            </div>
            <div>
              <p className="text-xs text-muted-foreground">
                Porcentaje legal
              </p>
              <p className="mt-1 text-lg font-semibold">
                {parameters.legalPercent}%
              </p>
            </div>
            <div>
              <p className="text-xs text-muted-foreground">
                Fondo total
              </p>
              <p className="mt-1 text-lg font-semibold text-primary">
                {money.format(calculateFund(parameters))}
              </p>
            </div>
          </div>
          <Separator />
          <div>
            <div className="mb-2 flex justify-between text-sm">
              <span>
                Distribuido
              </span>
              <span className="font-medium">
                {totals.fund > 0 ? ((totals.distributed / totals.fund) * 100).toFixed(1) : '0.0'}%
              </span>
            </div>
            <Progress
              value={totals.fund > 0 ? (totals.distributed / totals.fund) * 100 : 0}
              className="h-2"
            />
            <div className="mt-2 flex justify-between text-xs text-muted-foreground">
              <span>
                {money.format(totals.distributed)}
              </span>
              <span>
                Remanente {money.format(totals.remainder)}
              </span>
            </div>
          </div>
          <div className="rounded-lg border border-border/70 bg-muted/30 p-4">
            <div className="flex items-start gap-3">
              <div className="rounded-md bg-primary/10 p-2 text-primary">
                <Calculator className="size-4" />
              </div>
              <div>
                <p className="text-sm font-medium">
                  Fórmula de distribución
                </p>
                <p className="mt-1 text-xs leading-5 text-muted-foreground">
                  50% se asigna según días laborados y 50% según
                  remuneración computable. Cada resultado está sujeto
                  al tope de 18 remuneraciones.
                </p>
              </div>
            </div>
          </div>
        </div>
      </CardContent>
    </Card>
    {/* TOP 5 RESULTADOS */}
    <Card>
      <CardHeader>
        <CardTitle>
          Top 5 resultados
        </CardTitle>
        <CardDescription>
          Mayor monto individual calculado
        </CardDescription>
      </CardHeader>
      <CardContent>
        <div className="flex flex-col gap-4">
          {top.map((result, index) => (
            <button
              key={result.id}
              onClick={() => onSelect(result)}
              className="flex items-center gap-3 text-left"
            >
              <span className="flex size-6 items-center justify-center rounded-full bg-muted text-xs font-semibold">
                {index + 1}
             </span>
              <div className="min-w-0 flex-1">
                <p className="truncate text-sm font-medium">
                  {result.name}
                </p>
                <div className="mt-1 h-1.5 overflow-hidden rounded-full bg-muted">
                  <div
                    className="h-full rounded-full bg-primary"
                    style={{
                      width: `${Math.max(
                        8,
                        (result.finalAmount / top[0].finalAmount) * 100
                      )}%`,
                    }}
                  />
                </div>
              </div>
              <span className="text-sm font-semibold">
                {money.format(result.finalAmount)}
              </span>
            </button>
          ))}
        </div>
</CardContent>
</Card>
</div>
<div className="rounded-xl border border-border/60 bg-card overflow-hidden">
  <div className="border-b border-border/60 px-5 py-4">
    <p className="text-sm font-semibold">Detalle de distribución por trabajador</p>
    <p className="mt-1 text-xs text-muted-foreground">
      50% por días efectivos · 50% por remuneraciones computables · haz clic en una fila para ver la trazabilidad del cálculo
    </p>
  </div>
  <div className="overflow-x-auto">
    <Table>
      <TableHeader>
        <TableRow>
          <TableHead>DNI</TableHead>
          <TableHead>Trabajador</TableHead>
          <TableHead>Días efectivos</TableHead>
          <TableHead>% días</TableHead>
          <TableHead>Importe días</TableHead>
          <TableHead>Rem. computable</TableHead>
          <TableHead>% rem.</TableHead>
          <TableHead>Importe rem.</TableHead>
          <TableHead className="text-right">Utilidad bruta</TableHead>
        </TableRow>
      </TableHeader>
      <TableBody>
        {results.length === 0 && <TableRow><TableCell colSpan={9} className="py-8 text-center text-sm text-muted-foreground">Sin resultados calculados</TableCell></TableRow>}
        {results.map((result) => {
          const abierto = detalleDni === result.id
          const pctDias = totalDiasEfectivos > 0 ? (result.diasEfectivos / totalDiasEfectivos) * 100 : 0
          const pctRem = totalRemuneraciones > 0 ? (result.remuneracionComputable / totalRemuneraciones) * 100 : 0
          return (
            <Fragment key={result.id}>
              <TableRow className="cursor-pointer" onClick={() => setDetalleDni(abierto ? null : result.id)}>
                <TableCell>
                  <div className="flex items-center gap-2">
                    <ChevronRight className={`size-3.5 text-muted-foreground transition-transform ${abierto ? 'rotate-90' : ''}`} />
                    <span className="font-mono text-xs">{result.code}</span>
                  </div>
                </TableCell>
                <TableCell><p className="font-medium">{result.name}</p></TableCell>
                <TableCell>{number.format(result.diasEfectivos)}</TableCell>
                <TableCell>{totalDiasEfectivos > 0 ? `${pctDias.toFixed(4)}%` : '—'}</TableCell>
                <TableCell>{money.format(result.distribucion.utilidadPorDias)}</TableCell>
                <TableCell>{money.format(result.remuneracionComputable)}</TableCell>
                <TableCell>{totalRemuneraciones > 0 ? `${pctRem.toFixed(4)}%` : '—'}</TableCell>
                <TableCell>{money.format(result.distribucion.utilidadPorRemuneraciones)}</TableCell>
                <TableCell className="text-right font-semibold">{money.format(result.distribucion.utilidadBruta)}</TableCell>
              </TableRow>
              {abierto && (
                <TableRow>
                  <TableCell colSpan={9} className="bg-muted/30 p-0">
                    <div className="px-5 py-4">
                      <p className="mb-3 text-xs font-semibold uppercase tracking-[0.18em] text-primary">Trazabilidad del cálculo</p>
                      <div className="grid gap-4 md:grid-cols-2">
                        <div className="rounded-lg border border-border/70 bg-card p-4">
                          <p className="text-sm font-medium">Parte por días</p>
                          <dl className="mt-2 space-y-1 text-xs">
                            <div className="flex justify-between"><dt>Días efectivos</dt><dd>{number.format(result.diasEfectivos)}</dd></div>
                            <div className="flex justify-between"><dt>Total días efectivos</dt><dd>{number.format(totalDiasEfectivos)}</dd></div>
                            <div className="flex justify-between"><dt>Proporción</dt><dd>{totalDiasEfectivos > 0 ? `${number.format(result.diasEfectivos)} / ${number.format(totalDiasEfectivos)} = ${pctDias.toFixed(4)}%` : '—'}</dd></div>
                            <div className="flex justify-between"><dt>Fondo por días</dt><dd>{money.format(fondoPorDias)}</dd></div>
                            <div className="flex justify-between"><dt>Fondo × proporción</dt><dd>{totalDiasEfectivos > 0 ? `${money.format(fondoPorDias)} × ${pctDias.toFixed(4)}%` : '—'}</dd></div>
                            <div className="flex justify-between border-t border-border/70 pt-1.5"><dt className="font-medium">Importe por días</dt><dd className="font-semibold text-primary">{money.format(result.distribucion.utilidadPorDias)}</dd></div>
                          </dl>
                        </div>
                        <div className="rounded-lg border border-border/70 bg-card p-4">
                          <p className="text-sm font-medium">Parte por remuneraciones</p>
                          <dl className="mt-2 space-y-1 text-xs">
                            <div className="flex justify-between"><dt>Remuneración computable</dt><dd>{money.format(result.remuneracionComputable)}</dd></div>
                            <div className="flex justify-between"><dt>Total remuneraciones</dt><dd>{money.format(totalRemuneraciones)}</dd></div>
                            <div className="flex justify-between"><dt>Proporción</dt><dd>{totalRemuneraciones > 0 ? `${money.format(result.remuneracionComputable)} / ${money.format(totalRemuneraciones)} = ${pctRem.toFixed(4)}%` : '—'}</dd></div>
                            <div className="flex justify-between"><dt>Fondo por remuneraciones</dt><dd>{money.format(fondoPorRemuneraciones)}</dd></div>
                            <div className="flex justify-between"><dt>Fondo × proporción</dt><dd>{totalRemuneraciones > 0 ? `${money.format(fondoPorRemuneraciones)} × ${pctRem.toFixed(4)}%` : '—'}</dd></div>
                            <div className="flex justify-between border-t border-border/70 pt-1.5"><dt className="font-medium">Importe por remuneraciones</dt><dd className="font-semibold text-primary">{money.format(result.distribucion.utilidadPorRemuneraciones)}</dd></div>
                          </dl>
                        </div>
                      </div>
                      <div className="mt-4 rounded-lg bg-primary/5 p-4">
                        <p className="text-sm font-medium">Utilidad bruta</p>
                        <p className="mt-1 text-xs text-muted-foreground">
                          {money.format(result.distribucion.utilidadPorDias)} + {money.format(result.distribucion.utilidadPorRemuneraciones)}
                        </p>
                        <p className="mt-1 text-xl font-semibold text-primary">{money.format(result.distribucion.utilidadBruta)}</p>
                      </div>
                    </div>
                  </TableCell>
                </TableRow>
              )}
            </Fragment>
          )
        })}
        {results.length > 0 && (
          <TableRow className="bg-muted/40 font-semibold">
            <TableCell colSpan={2}>Totales</TableCell>
            <TableCell>{number.format(totalDiasEfectivos)}</TableCell>
            <TableCell>{totalDiasEfectivos > 0 ? '100%' : '—'}</TableCell>
            <TableCell>{money.format(sumaImporteDias)}</TableCell>
            <TableCell>{money.format(totalRemuneraciones)}</TableCell>
            <TableCell>{totalRemuneraciones > 0 ? '100%' : '—'}</TableCell>
            <TableCell>{money.format(sumaImporteRemuneraciones)}</TableCell>
            <TableCell className="text-right">{money.format(sumaUtilidadBruta)}</TableCell>
          </TableRow>
        )}
      </TableBody>
    </Table>
  </div>
</div>
  </div>
) : (
  <div className="flex flex-col gap-7">
    <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
      <div>
        <p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">
          Carga de información
        </p>
        <h2 className="text-xl font-semibold tracking-tight">
          Importar base de trabajadores
        </h2>
        <p className="mt-1 text-sm text-muted-foreground">
          Importa los archivos de remuneraciones y días laborados del ejercicio.
        </p>
      </div>
      <BotonExportarInformacion
        onClick={() => { void exportarCarga(trabajadores, trabajadoresDias) }}
        disabled={trabajadores.length === 0 && trabajadoresDias.length === 0}
      />
    </div>

    <Card className="border-primary/30 bg-primary/5">
    <CardContent className="flex flex-col gap-4 p-5 sm:flex-row sm:items-center">
    <div className="flex size-11 items-center justify-center rounded-xl bg-primary/15 text-primary"><Upload /></div>
  <div className="flex-1">
    <p className="font-semibold">Base de trabajadores lista para procesar</p>
  <p className="mt-1 text-sm text-muted-foreground">{listoParaProcesar ? `${trabajadoresUnificados.length} trabajadores consolidados (${trabajadores.length} remuneraciones · ${trabajadoresDias.length} días)` : 'Sin importaciones registradas — importa los archivos de remuneraciones y días laborados'}</p></div>
  <Button variant="outline" onClick={handleProcesarAhora} disabled={!listoParaProcesar}>Procesar ahora <ArrowRight data-icon="inline-end" /></Button></CardContent></Card>

  <div className="grid gap-6 xl:grid-cols-2">
  <ImportCard
    kind="remunerations"
    trabajadores={trabajadores}
    setTrabajadores={setTrabajadores}
  />
  <ImportCard
  kind="worked-days"
  trabajadores={trabajadoresDias}
  setTrabajadores={setTrabajadoresDias}
  setTrabajadoresDias={setTrabajadoresDias}
/>
  </div>
  </div>
)}
</div>
)
}
