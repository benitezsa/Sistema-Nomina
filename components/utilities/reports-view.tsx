'use client'

import { useState } from 'react'
import { Download, FileBarChart, FileSpreadsheet, Loader2, Printer, Users } from 'lucide-react'
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { money } from '@/lib/utilities-calculation'
import { useSettings } from '@/lib/settings-context'
import { buildConsolidadoHtml, buildIndividualHtml, buildMatrizHtml, exportConsolidado, exportIndividual, exportMatriz, openPrintWindow, type ReportContext } from '@/lib/report-export'
import type { EmployeeUtilityResult } from './types'

function SectionTitle({ eyebrow, title, description }: { eyebrow?: string; title: string; description?: string }) {
  return (
    <div>
      <p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">{eyebrow}</p>
      <h1 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">{title}</h1>
      {description && <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">{description}</p>}
    </div>
  )
}

type ReportType = 'consolidado' | 'individual' | 'matriz'

export function ReportsView({
  results,
  totals,
  onNotify,
}: {
  results: EmployeeUtilityResult[]
  totals: ReturnType<typeof import('@/lib/utilities-calculation').getTotals>
  onNotify: (m: string) => void
}) {
  const { company, parameters } = useSettings()
  const [busy, setBusy] = useState<ReportType | null>(null)

  const ctx: ReportContext = { company, parameters, results, totals }

  const handleDownload = async (type: ReportType) => {
    if (busy) return
    if ((type === 'individual' || type === 'matriz') && results.length === 0) {
      onNotify('No hay resultados de cálculo para generar este reporte')
      return
    }
    setBusy(type)
    try {
      if (type === 'consolidado') {
        await exportConsolidado(ctx)
        onNotify('Reporte consolidado descargado (Excel)')
      } else if (type === 'individual') {
        await exportIndividual(ctx)
        onNotify('Reporte individual descargado (Excel)')
      } else {
        await exportMatriz(ctx)
        onNotify('Matriz de cálculo descargada (Excel)')
      }
    } catch {
      onNotify('No se pudo generar el reporte')
    } finally {
      setBusy(null)
    }
  }

  const handlePrint = (type: ReportType) => {
    if (busy) return
    if ((type === 'individual' || type === 'matriz') && results.length === 0) {
      onNotify('No hay resultados de cálculo para imprimir este reporte')
      return
    }
    if (type === 'consolidado') {
      openPrintWindow('Reporte consolidado', buildConsolidadoHtml(ctx))
    } else if (type === 'individual') {
      openPrintWindow('Reporte individual', buildIndividualHtml(ctx))
    } else {
      openPrintWindow('Matriz de cálculo', buildMatrizHtml(ctx))
    }
    onNotify('Vista de impresión abierta')
  }

  const fecha = new Date().toLocaleDateString('es-PE', { day: '2-digit', month: 'short', year: 'numeric' })
  const reports: { type: ReportType; icon: typeof FileBarChart; title: string; description: string; meta: string }[] = [
    {
      type: 'consolidado',
      icon: FileBarChart,
      title: 'Reporte consolidado',
      description: 'Resumen del fondo, distribución y remanentes.',
      meta: results.length > 0 ? `${results.length} trabajadores · ${money.format(totals.distributed)} distribuido` : `Sin resultados · ${money.format(totals.fund)} fondo`,
    },
    {
      type: 'individual',
      icon: Users,
      title: 'Reporte individual',
      description: 'Detalle de cálculo por cada trabajador.',
      meta: results.length > 0 ? `Detalle de ${results.length} trabajadores · ${fecha}` : 'Sin resultados de cálculo',
    },
    {
      type: 'matriz',
      icon: FileSpreadsheet,
      title: 'Matriz de cálculo',
      description: 'Factores, topes y fórmulas trazables.',
      meta: `Formato Excel · ${fecha}`,
    },
  ]

  return (
    <div className="flex flex-col gap-7">
      <SectionTitle eyebrow="Centro de reportes" title="Reportes" description="Genera documentos listos para revisión, firma y archivo." />
      <div className="grid gap-4 lg:grid-cols-3">
        {reports.map(({ type, icon: Icon, title, description, meta }) => (
          <Card key={title}>
            <CardHeader>
              <div className="flex size-10 items-center justify-center rounded-lg bg-primary/8 text-primary">
                <Icon className="size-5" />
              </div>
              <CardTitle className="pt-2">{title}</CardTitle>
              <CardDescription>{description}</CardDescription>
            </CardHeader>
            <CardContent>
              <p className="text-xs text-muted-foreground">{meta}</p>
              <div className="mt-4 flex gap-2">
                <Button size="sm" disabled={busy !== null} onClick={() => handleDownload(type)}>
                  {busy === type ? <Loader2 className="mr-2 size-4 animate-spin" /> : <Download className="mr-2 size-4" />}Descargar
                </Button>
                <Button size="sm" variant="outline" disabled={busy !== null} onClick={() => handlePrint(type)}>
                  <Printer className="mr-2 size-4" />Imprimir
                </Button>
              </div>
            </CardContent>
          </Card>
        ))}
      </div>
      <Alert>
        <FileBarChart className="size-4" />
        <AlertTitle>Documentos auditables</AlertTitle>
        <AlertDescription>Cada reporte conserva la versión del cálculo, usuario responsable y fecha de generación.</AlertDescription>
      </Alert>
    </div>
  )
}