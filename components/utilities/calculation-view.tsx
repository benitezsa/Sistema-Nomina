'use client'

import { AlertCircle, ArrowDownToLine, ArrowRight, Calculator, Check, Download, FileSpreadsheet, Play, ShieldCheck, Upload, Users } from 'lucide-react'
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Progress } from '@/components/ui/progress'
import { Separator } from '@/components/ui/separator'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { calculateFund, getTotals, money, number } from '@/lib/utilities-calculation'
import { parameters } from '@/lib/utilities-data'
import type { EmployeeUtilityResult } from './types'

function SectionTitle({ eyebrow, title, description, action }: { eyebrow?: string; title: string; description?: string; action?: React.ReactNode }) { return <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between"><div><p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">{eyebrow}</p><h1 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">{title}</h1>{description && <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">{description}</p>}</div>{action}</div> }
function Kpi({ label, value, hint, tone = 'default', icon: Icon }: { label: string; value: string; hint: string; tone?: 'default' | 'teal' | 'amber' | 'rose'; icon: typeof Calculator }) { return <Card className="border-border/70 shadow-sm"><CardContent className="flex items-start justify-between p-5"><div><p className="text-xs font-medium text-muted-foreground">{label}</p><p className="mt-2 text-xl font-semibold tracking-tight">{value}</p><p className={`mt-1 text-xs ${tone === 'rose' ? 'text-rose-400' : tone === 'amber' ? 'text-amber-400' : tone === 'teal' ? 'text-teal-400' : 'text-muted-foreground'}`}>{hint}</p></div><div className={`rounded-lg p-2.5 ${tone === 'teal' ? 'bg-teal-500/10 text-teal-400' : tone === 'amber' ? 'bg-amber-500/10 text-amber-400' : tone === 'rose' ? 'bg-rose-500/10 text-rose-400' : 'bg-primary/8 text-primary'}`}><Icon className="size-4" /></div></CardContent></Card> }

  function ImportCard({ kind }: { kind: 'remunerations' | 'worked-days' }) {
  const isRemunerations = kind === 'remunerations'

  const title = isRemunerations
    ? 'Remuneraciones'
    : 'Días laborados'

  const description = isRemunerations
    ? 'Importa la información mensual o consolidada de remuneraciones computables.'
    : 'Carga los días trabajados, ausencias y licencias del ejercicio.'

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
        <button
          type="button"
          className="flex w-full flex-col items-center gap-3 rounded-2xl border border-dashed border-primary/40 bg-primary/5 px-6 py-9 text-center transition-colors hover:bg-primary/10"
        >
          <FileSpreadsheet className="size-8 text-primary" />

          <span className="font-medium">
            Arrastra el archivo aquí
          </span>

          <span className="text-sm text-muted-foreground">
            o selecciónalo desde tu dispositivo
          </span>

          <span className="rounded-lg border bg-background px-3 py-2 text-sm font-medium">
            Seleccionar archivo
          </span>
        </button>
      </CardContent>
    </Card>
  )
}
export function CalculationView({ results, totals, errors, onRun, onSelect }: { results: EmployeeUtilityResult[]; totals: ReturnType<typeof getTotals>; errors: string[]; onRun: () => void; onSelect: (result: EmployeeUtilityResult) => void }) { const top = [...results].sort((a,b) => b.finalAmount - a.finalAmount).slice(0,5); return <div className="flex flex-col gap-7"><SectionTitle eyebrow="Cálculo · Paso 2 de 3" title="Distribución de utilidades" description="Revisa los parámetros, valida los factores de distribución y genera el cálculo anual del ejercicio." action={<div className="flex gap-2 sm:hidden"><Button variant="outline" size="sm"><Check className="mr-2 size-4" />Guardar</Button><Button size="sm" onClick={onRun}><Play className="mr-2 size-4" />Ejecutar</Button></div>} /><Card className="border-primary/30 bg-primary/5"><CardContent className="flex flex-col gap-4 p-5 sm:flex-row sm:items-center"><div className="flex size-11 items-center justify-center rounded-xl bg-primary/15 text-primary"><Upload /></div><div className="flex-1"><p className="font-semibold">Base de trabajadores lista para procesar</p><p className="mt-1 text-sm text-muted-foreground">27 registros · Última importación 18 ago 2026 · Excel validado</p></div><Button variant="outline" onClick={onRun}>Procesar ahora <ArrowRight data-icon="inline-end" /></Button></CardContent></Card>
{/* Importación de información para el cálculo */}
<div className="grid gap-6 xl:grid-cols-2">
  <ImportCard kind="remunerations" />
  <ImportCard kind="worked-days" />
</div>

<div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
<Kpi label="Fondo legal" value={money.format(totals.fund)} hint="10% de la renta neta" icon={ShieldCheck} tone="teal" /><Kpi label="A distribuir" value={money.format(totals.distributed)} hint="50% días · 50% remuneración" icon={ArrowDownToLine} /><Kpi label="Trabajadores" value={number.format(results.length)} hint={`${results.length} registros procesados`} icon={Users} tone="amber" /><Kpi label="Remanente" value={money.format(totals.remainder)} hint={`${totals.capped} topes aplicados`} icon={AlertCircle} tone="rose" /></div>{errors.length > 0 && <Alert variant="destructive"><AlertCircle className="size-4" /><AlertTitle>Validaciones pendientes</AlertTitle><AlertDescription>{errors.join(' ')}</AlertDescription></Alert>}<div className="grid gap-5 xl:grid-cols-[1.15fr_.85fr]"><Card><CardHeader><div className="flex items-start justify-between gap-4"><div><CardTitle>Resumen del fondo</CardTitle><CardDescription>Distribución legal aplicable al ejercicio 2026</CardDescription></div><Badge variant="secondary" className="bg-emerald-500/10 text-emerald-400">Validado</Badge></div></CardHeader><CardContent><div className="flex flex-col gap-5"><div className="grid gap-4 sm:grid-cols-3"><div><p className="text-xs text-muted-foreground">Renta neta</p><p className="mt-1 text-lg font-semibold">{money.format(parameters.income)}</p></div><div><p className="text-xs text-muted-foreground">Porcentaje legal</p><p className="mt-1 text-lg font-semibold">{parameters.legalPercent}%</p></div><div><p className="text-xs text-muted-foreground">Fondo total</p><p className="mt-1 text-lg font-semibold text-primary">{money.format(calculateFund(parameters))}</p></div></div><Separator /><div><div className="mb-2 flex justify-between text-sm"><span>Distribuido</span><span className="font-medium">{((totals.distributed / totals.fund) * 100).toFixed(1)}%</span></div><Progress value={(totals.distributed / totals.fund) * 100} className="h-2" /><div className="mt-2 flex justify-between text-xs text-muted-foreground"><span>{money.format(totals.distributed)}</span><span>Remanente {money.format(totals.remainder)}</span></div></div><div className="rounded-lg border border-border/70 bg-muted/30 p-4"><div className="flex items-start gap-3"><div className="rounded-md bg-primary/10 p-2 text-primary"><Calculator className="size-4" /></div><div><p className="text-sm font-medium">Fórmula de distribución</p><p className="mt-1 text-xs leading-5 text-muted-foreground">50% se asigna según días laborados y 50% según remuneración computable. Cada resultado está sujeto al tope de 18 remuneraciones.</p></div></div></div></div></CardContent></Card><Card><CardHeader><CardTitle>Top 5 resultados</CardTitle><CardDescription>Mayor monto individual calculado</CardDescription></CardHeader><CardContent><div className="flex flex-col gap-4">{top.map((result, index) => <button key={result.id} onClick={() => onSelect(result)} className="flex items-center gap-3 text-left"><span className="flex size-6 items-center justify-center rounded-full bg-muted text-xs font-semibold">{index + 1}</span><div className="min-w-0 flex-1"><p className="truncate text-sm font-medium">{result.name}</p><div className="mt-1 h-1.5 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full bg-primary" style={{ width: `${Math.max(8, (result.finalAmount / top[0].finalAmount) * 100)}%` }} /></div></div><span className="text-sm font-semibold">{money.format(result.finalAmount)}</span></button>)}</div></CardContent></Card></div><Card><CardHeader><div className="flex items-center justify-between"><div><CardTitle>Distribución por trabajador</CardTitle><CardDescription>Vista previa de los resultados calculados</CardDescription></div><Button variant="outline" size="sm"><Download className="mr-2 size-4" />Exportar</Button></div></CardHeader><CardContent><div className="overflow-x-auto"><Table><TableHeader><TableRow><TableHead>Trabajador</TableHead><TableHead>Días</TableHead><TableHead>Remuneración</TableHead><TableHead className="text-right">Preliminar</TableHead><TableHead className="text-right">Final</TableHead></TableRow></TableHeader><TableBody>{results.slice(0,6).map((result) => <TableRow key={result.id} className="cursor-pointer" onClick={() => onSelect(result)}><TableCell><div><p className="font-medium">{result.name}</p><p className="text-xs text-muted-foreground">{result.code} · {result.department}</p></div></TableCell><TableCell>{number.format(result.days)}</TableCell><TableCell>{money.format(result.remuneration)}</TableCell><TableCell className="text-right">{money.format(result.preliminary)}</TableCell><TableCell className="text-right font-semibold text-primary">{money.format(result.finalAmount)}</TableCell></TableRow>)}</TableBody></Table></div></CardContent></Card></div> }
