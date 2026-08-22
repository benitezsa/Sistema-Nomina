'use client'

import { ArrowRight, BarChart3, Calculator, CheckCircle2, Clock3, FileSpreadsheet, Plus, TrendingUp, Upload, Users } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card'
import { Progress } from '@/components/ui/progress'
import type { CompanyClient, EmployeeUtilityResult } from './types'
const currency = new Intl.NumberFormat('es-PE', { style: 'currency', currency: 'PEN' })

type Props = { companies: CompanyClient[]; results: EmployeeUtilityResult[]; onNavigate: (view: 'Empresas' | 'Trabajadores' | 'Cálculo' | 'Resultados') => void }

const tasks = [
  { title: 'Completar datos de trabajadores', meta: '2 registros observados', progress: 84, tone: 'amber' },
  { title: 'Revisar parámetros legales', meta: 'Ejercicio 2026', progress: 100, tone: 'teal' },
  { title: 'Generar reporte final', meta: 'Pendiente de cálculo', progress: 35, tone: 'indigo' },
]

export function DashboardView({ companies, results, onNavigate }: Props) {
  const total = results.reduce((sum, row) => sum + row.finalAmount, 0)
  const activeCompanies = companies.filter((company) => company.status === 'Activo').length
  const observed = results.filter((row) => row.status !== 'Completo').length

  return <div className="flex flex-col gap-6">
    <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
      <div><p className="text-sm text-muted-foreground">Lunes, 18 de agosto de 2026</p><h1 className="mt-1 text-3xl font-semibold tracking-tight">Buenos días, Mariana</h1><p className="mt-2 text-muted-foreground">Este es el estado de tus operaciones de nómina.</p></div>
      <div className="flex flex-wrap gap-2"><Button variant="outline" onClick={() => onNavigate('Empresas')}><Plus data-icon="inline-start" />Nueva empresa</Button><Button onClick={() => onNavigate('Cálculo')}><Calculator data-icon="inline-start" />Nuevo cálculo</Button></div>
    </div>
    <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
      <Metric label="Empresas activas" value={String(activeCompanies)} detail="3 con ejercicio abierto" icon={BarChart3} />
      <Metric label="Trabajadores" value={String(results.length)} detail="Base actualizada" icon={Users} />
      <Metric label="Por distribuir" value={currency.format(total)} detail="Ejercicio 2026" icon={TrendingUp} accent />
      <Metric label="Tareas pendientes" value={String(observed + 2)} detail="Requieren atención" icon={Clock3} warning />
    </div>
    <div className="grid gap-6 xl:grid-cols-[1.35fr_0.65fr]">
      <Card><CardHeader className="flex flex-row items-center justify-between"><div><CardTitle>Bandeja de trabajo</CardTitle><p className="mt-1 text-sm text-muted-foreground">Prioridades de tu equipo</p></div><Badge variant="secondary">Esta semana</Badge></CardHeader><CardContent className="flex flex-col gap-3">{tasks.map((task) => <button key={task.title} className="group flex items-center gap-4 rounded-lg border border-border/70 bg-muted/20 p-4 text-left transition-colors hover:bg-accent" onClick={() => onNavigate('Cálculo')}><div className="flex size-10 shrink-0 items-center justify-center rounded-lg bg-background text-primary"><FileSpreadsheet /></div><div className="min-w-0 flex-1"><div className="flex items-center justify-between gap-3"><p className="truncate font-medium">{task.title}</p><span className="text-xs text-muted-foreground">{task.progress}%</span></div><p className="mt-1 text-xs text-muted-foreground">{task.meta}</p><Progress value={task.progress} className="mt-3 h-1.5" /></div><ArrowRight className="text-muted-foreground transition-transform group-hover:translate-x-1" /></button>)}</CardContent></Card>
      <Card><CardHeader><CardTitle>Actividad reciente</CardTitle><p className="mt-1 text-sm text-muted-foreground">Últimos movimientos</p></CardHeader><CardContent className="flex flex-col gap-5">{['Cálculo 2026 ejecutado', 'Empresa Andinas actualizada', 'Archivo de trabajadores importado'].map((item, index) => <div key={item} className="flex gap-3"><div className="mt-0.5 flex size-7 shrink-0 items-center justify-center rounded-full bg-primary/15 text-primary"><CheckCircle2 /></div><div><p className="text-sm font-medium">{item}</p><p className="mt-1 text-xs text-muted-foreground">{index + 1}h atrás · NóminaPro</p></div></div>)}</CardContent></Card>
    </div>
    <Card><CardHeader><CardTitle>Acciones rápidas</CardTitle><p className="mt-1 text-sm text-muted-foreground">Continúa donde lo dejaste</p></CardHeader><CardContent className="grid gap-3 sm:grid-cols-3"><QuickAction icon={Upload} title="Importar trabajadores" onClick={() => onNavigate('Trabajadores')} /><QuickAction icon={Calculator} title="Procesar utilidades" onClick={() => onNavigate('Cálculo')} /><QuickAction icon={BarChart3} title="Ver resultados" onClick={() => onNavigate('Resultados')} /></CardContent></Card>
  </div>
}

function Metric({ label, value, detail, icon: Icon, accent, warning }: { label: string; value: string; detail: string; icon: typeof Users; accent?: boolean; warning?: boolean }) { return <Card><CardContent className="flex items-start justify-between p-5"><div><p className="text-sm text-muted-foreground">{label}</p><p className={`mt-2 text-2xl font-semibold ${accent ? 'text-primary' : warning ? 'text-amber-400' : ''}`}>{value}</p><p className="mt-1 text-xs text-muted-foreground">{detail}</p></div><div className="flex size-10 items-center justify-center rounded-lg bg-primary/12 text-primary"><Icon /></div></CardContent></Card> }
function QuickAction({ icon: Icon, title, onClick }: { icon: typeof Upload; title: string; onClick: () => void }) { return <Button variant="outline" className="h-auto justify-start gap-3 p-4" onClick={onClick}><Icon data-icon="inline-start" /><span>{title}</span><ArrowRight className="ml-auto" /></Button> }
