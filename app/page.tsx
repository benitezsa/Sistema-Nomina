'use client'

import { useEffect, useMemo, useState } from 'react'
import { ArrowDownToLine, ArrowRight, BarChart3, Bell, BookOpen, Calculator, Check, ChevronDown, FileBarChart, History, LayoutDashboard, Menu, MoreHorizontal, Play, Plus, Settings2, SlidersHorizontal, Sparkles, Table2, Users, X } from 'lucide-react'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Card, CardContent } from '@/components/ui/card'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog'
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from '@/components/ui/sheet'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip'
import { company, companyClients, employees, exercises, parameters } from '@/lib/utilities-data'
import { calculateResults, getTotals, money, number, validateParameters } from '@/lib/utilities-calculation'
import { AuditView } from '@/components/utilities/audit-view'
import { CalculationView } from '@/components/utilities/calculation-view'
import { CompaniesView } from '@/components/utilities/companies-view'
import { ConfigurationView } from '@/components/utilities/configuration-view'
import { DashboardView } from '@/components/utilities/dashboard-view'
import { RemaindersView } from '@/components/utilities/remainders-view'
import { ReportsView } from '@/components/utilities/reports-view'
import { ResultsView } from '@/components/utilities/results-view'
import { WorkersView } from '@/components/utilities/workers-view'
import type { EmployeeUtilityResult, ViewKey } from '@/components/utilities/types'

const allNavItems: { label: ViewKey | 'Dashboard'; icon: typeof LayoutDashboard; count?: string }[] = [
  { label: 'Dashboard', icon: LayoutDashboard },
  { label: 'Empresas', icon: LayoutDashboard },
  { label: 'Configuración', icon: Settings2 },
  { label: 'Trabajadores', icon: Users, count: '27' },
  { label: 'Cálculo', icon: Calculator },
  { label: 'Resultados', icon: BarChart3 },
  { label: 'Remanentes', icon: ArrowDownToLine },
  { label: 'Reportes', icon: FileBarChart },
  { label: 'Auditoría', icon: History },
]
const statusStyle: Record<string, string> = { Calculado: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30', Cerrado: 'bg-slate-500/10 text-slate-300 border-slate-500/30', 'En cálculo': 'bg-amber-500/10 text-amber-400 border-amber-500/30', Borrador: 'bg-blue-500/10 text-blue-400 border-blue-500/30', Completo: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30', Pendiente: 'bg-amber-500/10 text-amber-400 border-amber-500/30', Observado: 'bg-rose-500/10 text-rose-400 border-rose-500/30' }

function Status({ value }: { value: string }) { return <Badge variant="outline" className={`font-medium ${statusStyle[value] || ''}`}>{value}</Badge> }
function SectionTitle({ eyebrow, title, description, action }: { eyebrow?: string; title: string; description?: string; action?: React.ReactNode }) { return <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between"><div><p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">{eyebrow}</p><h1 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">{title}</h1>{description && <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">{description}</p>}</div>{action}</div> }
function Kpi({ label, value, hint, tone = 'default', icon: Icon }: { label: string; value: string; hint: string; tone?: 'default' | 'teal' | 'amber' | 'rose'; icon: typeof Calculator }) { return <Card className="border-border/70 shadow-sm"><CardContent className="flex items-start justify-between p-5"><div><p className="text-xs font-medium text-muted-foreground">{label}</p><p className="mt-2 text-xl font-semibold tracking-tight">{value}</p><p className={`mt-1 text-xs ${tone === 'rose' ? 'text-rose-400' : tone === 'amber' ? 'text-amber-400' : tone === 'teal' ? 'text-teal-400' : 'text-muted-foreground'}`}>{hint}</p></div><div className={`rounded-lg p-2.5 ${tone === 'teal' ? 'bg-teal-500/10 text-teal-400' : tone === 'amber' ? 'bg-amber-500/10 text-amber-400' : tone === 'rose' ? 'bg-rose-500/10 text-rose-400' : 'bg-primary/8 text-primary'}`}><Icon className="size-4" /></div></CardContent></Card> }

export default function Page() {
  const [view, setView] = useState<ViewKey | 'Dashboard'>('Dashboard')
  const [mobileNav, setMobileNav] = useState(false)
  const [showAllModules, setShowAllModules] = useState(false)
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState<EmployeeUtilityResult | null>(null)
  const [runOpen, setRunOpen] = useState(false)
  const [toast, setToast] = useState('')
  const results = useMemo(() => calculateResults(employees, parameters), [])
  const totals = useMemo(() => getTotals(results, parameters), [results])
  const validationErrors = validateParameters(parameters, employees.length)
  const filteredEmployees = employees.filter((e) => `${e.name} ${e.code} ${e.department}`.toLowerCase().includes(query.toLowerCase()))
  const notify = (message: string) => { setToast(message); window.setTimeout(() => setToast(''), 2600) }
  const go = (next: ViewKey | 'Dashboard') => { setView(next); setMobileNav(false) }
  const navItems = showAllModules
  ? allNavItems
  : allNavItems.filter(
      (item) => item.label === 'Dashboard' || item.label === 'Empresas'
    )
  useEffect(() => {
  const modulesUnlocked = localStorage.getItem('utilities-modules-unlocked')

  if (modulesUnlocked === 'true') {
    setShowAllModules(true)
  }
}, [])

  return <TooltipProvider><div className="min-h-screen bg-background text-foreground"><aside className={`fixed inset-y-0 left-0 z-40 flex w-64 flex-col border-r border-sidebar-border bg-sidebar transition-transform lg:translate-x-0 ${mobileNav ? 'translate-x-0' : '-translate-x-full'}`}><div className="flex h-16 items-center gap-3 border-b border-sidebar-border px-5"><div className="flex size-8 items-center justify-center rounded-lg bg-primary text-primary-foreground"><Sparkles className="size-4" /></div><div><p className="text-sm font-semibold">NóminaPro</p><p className="text-[10px] uppercase tracking-widest text-sidebar-foreground/50">Enterprise</p></div><Button variant="ghost" size="icon" className="ml-auto lg:hidden" onClick={() => setMobileNav(false)}><X /></Button></div><div className="flex flex-1 flex-col gap-6 p-3"><div><p className="px-3 pb-2 text-[10px] font-semibold uppercase tracking-[0.18em] text-sidebar-foreground/45">Módulo activo</p><div className="rounded-lg border border-sidebar-border bg-sidebar-accent/70 p-3"><div className="flex items-center gap-2"><Calculator className="size-4 text-primary" /><span className="text-sm font-medium">Utilidades</span></div><p className="mt-1 pl-6 text-xs text-sidebar-foreground/55">Cálculo y distribución</p></div></div><nav className="flex flex-col gap-1" aria-label="Navegación principal">{navItems.map(({ label, icon: Icon, count }) => <button key={label} onClick={() => go(label)} className={`flex items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm transition-colors ${view === label ? 'bg-primary text-primary-foreground shadow-sm' : 'text-sidebar-foreground/70 hover:bg-sidebar-accent hover:text-sidebar-foreground'}`}><Icon className="size-4" /><span className="flex-1">{label}</span>{count && <span className={`rounded px-1.5 py-0.5 text-[10px] ${view === label ? 'bg-primary-foreground/15' : 'bg-sidebar-accent'}`}>{count}</span>}</button>)}</nav></div><div className="border-t border-sidebar-border p-3"><div className="flex items-center gap-3 rounded-lg p-2"><Avatar className="size-8"><AvatarFallback className="bg-primary/10 text-xs text-primary">MF</AvatarFallback></Avatar><div className="min-w-0 flex-1"><p className="truncate text-xs font-medium">Mariana Flores</p><p className="truncate text-[11px] text-sidebar-foreground/50">Administrador</p></div><MoreHorizontal className="size-4 text-sidebar-foreground/40" /></div></div></aside>{mobileNav && <button aria-label="Cerrar menú" className="fixed inset-0 z-30 bg-foreground/20 lg:hidden" onClick={() => setMobileNav(false)} />}
<div className="lg:pl-64"><header className="sticky top-0 z-20 flex h-16 items-center gap-3 border-b border-border/70 bg-background/90 px-4 backdrop-blur sm:px-6"><Button variant="ghost" size="icon" className="lg:hidden" onClick={() => setMobileNav(true)}><Menu /></Button><div className="min-w-0 flex-1"><div className="flex items-center gap-2 text-xs text-muted-foreground"><span className="hidden sm:inline">Utilidades</span><span className="hidden sm:inline">/</span><span className="font-medium text-foreground">Ejercicio 2026</span><Status value="Calculado" /></div><p className="mt-0.5 truncate text-xs text-muted-foreground">{company.name} · RUC {company.ruc}</p></div><Tooltip><TooltipTrigger render={<Button variant="ghost" size="icon" />}><Bell className="size-4" /></TooltipTrigger><TooltipContent>Notificaciones</TooltipContent></Tooltip><Button variant="outline" size="sm" className="hidden sm:flex" onClick={() => notify('Centro de ayuda abierto')}><BookOpen className="mr-2 size-4" />Ayuda</Button></header>
      <main className="mx-auto max-w-[1600px] p-4 sm:p-6 lg:p-8"><div className="mb-6 flex items-center justify-between gap-4"><div><p className="text-xs font-medium text-muted-foreground">Marzo 2026</p><p className="mt-1 text-sm font-medium">Cierre anual de utilidades</p></div><div className="hidden items-center gap-2 md:flex"><div className="flex items-center gap-2 text-xs text-muted-foreground"><span className="flex size-6 items-center justify-center rounded-full bg-primary text-primary-foreground">1</span><span>Información</span></div><div className="h-px w-8 bg-border" /><div className="flex items-center gap-2 text-xs text-primary"><span className="flex size-6 items-center justify-center rounded-full bg-primary text-primary-foreground">2</span><span className="font-medium">Cálculo</span></div><div className="h-px w-8 bg-border" /><div className="flex items-center gap-2 text-xs text-muted-foreground"><span className="flex size-6 items-center justify-center rounded-full border border-border">3</span><span>Revisión</span></div></div><div className="hidden gap-2 sm:flex"><Button variant="outline" size="sm" onClick={() => notify('Cambios guardados en borrador')}><Check className="mr-2 size-4" />Guardar borrador</Button><Button size="sm" onClick={() => setRunOpen(true)}><Play className="mr-2 size-4" />Ejecutar cálculo</Button></div></div>
      {view === 'Dashboard' && <DashboardView companies={companyClients} results={results} onNavigate={(next) => go(next)} />}{view === 'Cálculo' && <CalculationView results={results} totals={totals} errors={validationErrors} onRun={() => setRunOpen(true)} onSelect={setSelected} />}
      {view === 'Empresas' && (
  <CompaniesView
    onOpen={() => go('Cálculo')}
    onNotify={notify}
    onShowAllModules={() => {
      setShowAllModules(true)
      go('Configuración')
    }}
  />
)}
      {view === 'Configuración' && <ConfigurationView onNotify={notify} />}
      {view === 'Trabajadores' && <WorkersView query={query} setQuery={setQuery} employees={filteredEmployees} onNotify={notify} />}
      {view === 'Resultados' && <ResultsView results={results} totals={totals} onSelect={setSelected} onNotify={notify} />}
      {view === 'Remanentes' && <RemaindersView results={results} totals={totals} />}
      {view === 'Reportes' && <ReportsView onNotify={notify} />}
      {view === 'Auditoría' && <AuditView onNotify={notify} />}
      </main></div>
      <Sheet open={!!selected} onOpenChange={(open) => !open && setSelected(null)}><SheetContent className="w-full sm:max-w-lg"><SheetHeader><SheetTitle>Detalle individual</SheetTitle><SheetDescription>Resultado calculado para el ejercicio 2026.</SheetDescription></SheetHeader>{selected && <div className="flex flex-col gap-6 p-5"><div className="flex items-center gap-3"><Avatar className="size-12"><AvatarFallback className="bg-primary/10 text-primary">{selected.name.split(' ').map((n) => n[0]).slice(0,2).join('')}</AvatarFallback></Avatar><div><p className="font-semibold">{selected.name}</p><p className="text-sm text-muted-foreground">{selected.role} · {selected.code}</p></div></div><div className="grid grid-cols-2 gap-3">{[['Días laborados', number.format(selected.days)],['Remuneración', money.format(selected.remuneration)],['Factor días', `${(selected.daysFactor * 100).toFixed(2)}%`],['Factor remuneración', `${(selected.remunerationFactor * 100).toFixed(2)}%`]].map(([l,v]) => <div key={l} className="rounded-lg border border-border/70 p-3"><p className="text-xs text-muted-foreground">{l}</p><p className="mt-1 font-semibold">{v}</p></div>)}</div><div className="rounded-xl bg-primary/5 p-4"><p className="text-sm text-muted-foreground">Monto final a distribuir</p><p className="mt-1 text-3xl font-semibold text-primary">{money.format(selected.finalAmount)}</p><div className="mt-3 flex justify-between text-xs"><span>Preliminar</span><span>{money.format(selected.preliminary)}</span></div><div className="mt-1 flex justify-between text-xs"><span>Tope ({parameters.capMonths} remuneraciones)</span><span>{money.format(selected.cap)}</span></div></div></div>}</SheetContent></Sheet>
      <Dialog open={runOpen} onOpenChange={setRunOpen}><DialogContent><DialogHeader><DialogTitle>Ejecutar cálculo de utilidades</DialogTitle><DialogDescription>Se generará una nueva versión con la información validada del ejercicio.</DialogDescription></DialogHeader><div className="flex flex-col gap-3 py-2"><div className="flex items-center gap-3 rounded-lg border border-emerald-500/30 bg-emerald-500/10 p-3 text-sm text-emerald-400"><Check className="size-4" />27 trabajadores listos para procesar</div><div className="flex items-center gap-3 rounded-lg border border-emerald-500/30 bg-emerald-500/10 p-3 text-sm text-emerald-400"><Check className="size-4" />Parámetros y fórmulas validados</div><p className="text-xs text-muted-foreground">El cálculo reemplazará la versión 1.2 solo al confirmar.</p></div><DialogFooter><Button variant="outline" onClick={() => setRunOpen(false)}>Cancelar</Button><Button onClick={() => { setRunOpen(false); notify('Cálculo ejecutado correctamente · versión 1.3') }}><Play className="mr-2 size-4" />Confirmar ejecución</Button></DialogFooter></DialogContent></Dialog>
      {toast && <div role="status" className="fixed bottom-5 right-5 z-50 flex items-center gap-3 rounded-lg border border-border bg-card px-4 py-3 text-sm shadow-lg"><Check className="size-4 text-emerald-400" />{toast}</div>}
    </div></TooltipProvider>
}
function LegacyExercisesView({ onOpen, onNotify }: { onOpen: () => void; onNotify: (m: string) => void }) { return <div className="flex flex-col gap-7"><SectionTitle eyebrow="Gestión de ejercicios" title="Ejercicios de utilidades" description="Administra los periodos anuales de cálculo y consulta su estado de cierre." action={<Button onClick={() => onNotify('Formulario de nuevo ejercicio abierto')}><Plus className="mr-2 size-4" />Nuevo ejercicio</Button>} /><Card><CardContent className="p-0"><div className="overflow-x-auto"><Table><TableHeader><TableRow><TableHead>Ejercicio</TableHead><TableHead>Empresa</TableHead><TableHead>Trabajadores</TableHead><TableHead>Fondo legal</TableHead><TableHead>Estado</TableHead><TableHead className="text-right">Acción</TableHead></TableRow></TableHeader><TableBody>{exercises.map((exercise) => <TableRow key={exercise.id}><TableCell><p className="font-semibold">{exercise.year}</p><p className="text-xs text-muted-foreground">Actualizado {exercise.updatedAt}</p></TableCell><TableCell><p className="font-medium">{exercise.company.name}</p><p className="text-xs text-muted-foreground">{exercise.company.ruc}</p></TableCell><TableCell>{exercise.employeeCount}</TableCell><TableCell>{money.format(exercise.fund)}</TableCell><TableCell><Status value={exercise.status} /></TableCell><TableCell className="text-right"><Button variant="outline" size="sm" onClick={onOpen}>Abrir <ArrowRight className="ml-2 size-4" /></Button></TableCell></TableRow>)}</TableBody></Table></div></CardContent></Card></div> }
