'use client'

import { useState } from 'react'
import { AlertCircle, ArrowDownToLine, BarChart3, Calculator, Download, FileText, Loader2 } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { getTotals, money, number } from '@/lib/utilities-calculation'
import { useSettings } from '@/lib/settings-context'
import { exportBoletas, exportResultados, type ReportContext } from '@/lib/report-export'
import type { EmployeeUtilityResult } from './types'

function SectionTitle({ eyebrow, title, description, action }: { eyebrow?: string; title: string; description?: string; action?: React.ReactNode }) { return <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
    <div>
        <p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">{eyebrow}</p>
        <h1 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">{title}</h1>{description && <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">{description}</p>}</div>{action}</div> }
function Kpi({ label, value, hint, tone = 'default', icon: Icon }: { label: string; value: string; hint: string; tone?: 'default' | 'teal' | 'amber' | 'rose'; icon: typeof Calculator }) { return <Card className="border-border/70 shadow-sm">
    <CardContent className="flex items-start justify-between p-5">
        <div>
        <p className="text-xs font-medium text-muted-foreground">{label}</p>
        <p className="mt-2 text-xl font-semibold tracking-tight">{value}</p>
<p className={`mt-1 text-xs ${tone === 'rose' ? 'text-rose-400' : tone === 'amber' ? 'text-amber-400' : tone === 'teal' ? 'text-teal-400' : 'text-muted-foreground'}`}>{hint}</p>
</div>
<div className={`rounded-lg p-2.5 ${tone === 'teal' ? 'bg-teal-500/10 text-teal-400' : tone === 'amber' ? 'bg-amber-500/10 text-amber-400' : tone === 'rose' ? 'bg-rose-500/10 text-rose-400' : 'bg-primary/8 text-primary'}`}>
<Icon className="size-4" />
</div>
</CardContent>
</Card> }

export function ResultsView({ results, totals, onSelect, onNotify, onConstancia }: { results: EmployeeUtilityResult[]; totals: ReturnType<typeof getTotals>; onSelect: (r: EmployeeUtilityResult) => void; onNotify: (m: string) => void; onConstancia: (r: EmployeeUtilityResult) => void }) {
  const { company, parameters } = useSettings()
  const [busy, setBusy] = useState<'boletas' | 'excel' | null>(null)
  const sorted = [...results].sort((a, b) => b.finalAmount - a.finalAmount)
  const ctx: ReportContext = { company, parameters, results, totals }

  const handleBoletas = async () => {
    if (busy || results.length === 0) return
    setBusy('boletas')
    try {
      await exportBoletas(ctx)
      onNotify('Boletas individuales descargadas (Excel)')
    } catch {
      onNotify('No se pudieron generar las boletas')
    } finally {
      setBusy(null)
    }
  }

  const handleExcel = async () => {
    if (busy) return
    setBusy('excel')
    try {
      await exportResultados(ctx)
      onNotify('Reporte de resultados descargado (Excel)')
    } catch {
      onNotify('No se pudo exportar el reporte')
    } finally {
      setBusy(null)
    }
  }

  return <div className="flex flex-col gap-7">
    <SectionTitle eyebrow="Resultados finales" title="Resultados de distribución" description="Consulta, valida y exporta el resultado final por trabajador, incluida la quinta categoría." action={<div className="flex gap-2">
      <Button onClick={handleBoletas} disabled={busy !== null || sorted.length === 0}>
        {busy === 'boletas' ? <Loader2 className="mr-2 size-4 animate-spin" /> : <FileText className="mr-2 size-4" />}Boleta individual</Button>
      <Button variant="outline" onClick={handleExcel} disabled={busy !== null}>
        {busy === 'excel' ? <Loader2 className="mr-2 size-4 animate-spin" /> : <Download className="mr-2 size-4" />}Exportar resultados</Button>
      </div>} />
<div className="grid gap-4 sm:grid-cols-4"><Kpi label="Total distribuido" value={money.format(totals.distributed)} hint="Conciliado contra el fondo" icon={ArrowDownToLine} tone="teal" />
<Kpi label="Quinta categoría total" value={money.format(totals.totalQuintaCategoria)} hint="Retención del ejercicio" icon={BarChart3} tone="amber" />
    <Kpi label="Promedio individual" value={money.format(results.length > 0 ? totals.distributed / results.length : 0)} hint="Por trabajador" icon={Calculator} />
    <Kpi label="Topes aplicados" value={number.format(totals.capped)} hint="Excedente trasladado a remanente" icon={AlertCircle} tone="rose" /></div>
<Card>
    <CardHeader>
        <CardTitle>Ranking de resultados</CardTitle>
<CardDescription>Detalle por trabajador: días efectivos, remuneración computable y utilidad neta.</CardDescription>
</CardHeader>
<CardContent>
    <div className="overflow-x-auto"><Table>
        <TableHeader>
            <TableRow>
                <TableHead>#</TableHead>
                <TableHead>Trabajador</TableHead>
                <TableHead>Días efectivos</TableHead>
                <TableHead>Rem. computable</TableHead>
                <TableHead>Util. por días</TableHead>
                <TableHead>Util. por rem.</TableHead>
                <TableHead>Utilidad bruta</TableHead>
                <TableHead>Quinta cat.</TableHead>
                <TableHead className="text-right">Utilidad neta</TableHead>
            </TableRow>
        </TableHeader>
                <TableBody>{sorted.map((r, i) => <TableRow key={r.id} className="cursor-pointer" onClick={() => onSelect(r)}>
                    <TableCell className="text-muted-foreground">{i+1}</TableCell>
                    <TableCell><p className="font-medium">{r.name}</p>
                    <p className="text-xs text-muted-foreground">{r.department}</p>
                    </TableCell>
                    <TableCell>{number.format(r.diasEfectivos)}</TableCell>
                    <TableCell>{money.format(r.remuneracionComputable)}</TableCell>
                    <TableCell>{money.format(r.distribucion.utilidadPorDias)}</TableCell>
                    <TableCell>{money.format(r.distribucion.utilidadPorRemuneraciones)}</TableCell>
                    <TableCell>{money.format(r.distribucion.utilidadBruta)}</TableCell>
                    <TableCell>{money.format(r.distribucion.quintaCategoria)}</TableCell>
                    <TableCell className="text-right font-semibold text-primary">{money.format(r.distribucion.utilidadNeta)} {r.capApplied && <Badge variant="outline" className="ml-2 bg-amber-500/10 text-amber-400">Tope</Badge>}</TableCell>
                    </TableRow>)}</TableBody>
</Table>
</div>
</CardContent>
</Card>
</div> }