'use client'

import { useEffect, useMemo, useState } from 'react'
import { useRouter } from 'next/navigation'
import { ArrowDownToLine, BarChart3, Bell, BookOpen, Calculator, CalendarDays, Check, FileBarChart, FileText, History, Info, LayoutDashboard, LogOut, Menu, Play, Settings2, Sparkles, Users, Wallet, X } from 'lucide-react'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Button } from '@/components/ui/button'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog'
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from '@/components/ui/sheet'
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip'
import { companyClients, employees } from '@/lib/utilities-data'
import { SettingsProvider, useSettings } from '@/lib/settings-context'
import { useAuth } from '@/lib/auth-context'
import { calculateResults, getTotals, money, number, validateParameters } from '@/lib/utilities-calculation'
import { AuditView } from '@/components/utilities/audit-view'
import { BoletaConstancia } from '@/components/utilities/boleta-view'
import { CalculationView } from '@/components/utilities/calculation-view'
import { CompaniesView } from '@/components/utilities/companies-view'
import { ConfigurationView } from '@/components/utilities/configuration-view'
import { DashboardView } from '@/components/utilities/dashboard-view'
import { RemaindersView } from '@/components/utilities/remainders-view'
import { ReportsView } from '@/components/utilities/reports-view'
import { ResultsView } from '@/components/utilities/results-view'
import { WorkersView } from '@/components/utilities/workers-view'
import type { EmployeeUtilityResult, ImportSummary, ViewKey } from '@/components/utilities/types'

const allNavItems: { label: ViewKey | 'Dashboard'; icon: typeof LayoutDashboard; count?: string }[] = [
  { label: 'Dashboard', icon: LayoutDashboard },
  { label: 'Empresas', icon: LayoutDashboard },
  { label: 'Configuración', icon: Settings2 },
  { label: 'Trabajadores', icon: Users },
  { label: 'Cálculo', icon: Calculator },
  { label: 'Resultados', icon: BarChart3 },
  { label: 'Remanentes', icon: ArrowDownToLine },
  { label: 'Reportes', icon: FileBarChart },
  { label: 'Auditoría', icon: History },
]
export default function Page() {
  return (
    <SettingsProvider>
      <App />
    </SettingsProvider>
  )
}

function App() {
  const { company, parameters } = useSettings()
  const { user, cargando, logout } = useAuth()
  const router = useRouter()
  const [view, setView] = useState<ViewKey | 'Dashboard'>('Dashboard')
  const [mobileNav, setMobileNav] = useState(false)
  const [query, setQuery] = useState('')
  const [selected, setSelected] = useState<EmployeeUtilityResult | null>(null)
  const [constancia, setConstancia] = useState<EmployeeUtilityResult | null>(null)
  const [runOpen, setRunOpen] = useState(false)
  const [toast, setToast] = useState('')
  const [importSummary, setImportSummary] = useState<ImportSummary>({
    remuneracionesTrabajadores: 0,
    totalRemuneracionAnual: 0,
    diasTrabajadores: 0,
    totalDiasRegistrados: 0,
    totalIncidencias: 0,
  })
  const [processedEmployees, setProcessedEmployees] = useState<typeof employees>(employees)
  const results = useMemo(() => calculateResults(processedEmployees, parameters), [processedEmployees, parameters])
  const totals = useMemo(() => getTotals(results, parameters), [results, parameters])
  const validationErrors = validateParameters(parameters, processedEmployees.length)
  const filteredEmployees = processedEmployees.filter((e) => `${e.name} ${e.code} ${e.department}`.toLowerCase().includes(query.toLowerCase()))
  const handleReprocessed = (rows: typeof employees) => { setProcessedEmployees(rows.map(r => ({ ...r, remuneration: r.remuneration }))); notify('Días y resultados reprocesados correctamente') }
  const notify = (message: string) => { setToast(message); window.setTimeout(() => setToast(''), 2600) }
  const go = (next: ViewKey | 'Dashboard') => { setView(next); setMobileNav(false) }
  const navItems = allNavItems

  useEffect(() => {
    if (!cargando && !user) router.push('/sign-in')
  }, [cargando, user, router])
  if (cargando) return <div className="min-h-screen bg-background" />
  if (!user) return null

  return <TooltipProvider>
    <div className="min-h-screen bg-background text-foreground">
      <aside className={`fixed inset-y-0 left-0 z-40 flex w-64 flex-col border-r border-sidebar-border bg-sidebar transition-transform lg:translate-x-0 ${mobileNav ? 'translate-x-0' : '-translate-x-full'}`}>
        <div className="flex h-16 items-center gap-3 border-b border-sidebar-border px-5">
          <div className="flex size-8 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <Sparkles className="size-4" />
            </div>
            <div>
              <p className="text-sm font-semibold">Utilidades</p>
              </div>
              <Button variant="ghost" size="icon" className="ml-auto lg:hidden" onClick={() => setMobileNav(false)}>
                <X />
                </Button>
                </div>
                <div className="flex flex-1 flex-col gap-6 p-3">
                  <div>
                    <p className="px-3 pb-2 text-[10px] font-semibold uppercase tracking-[0.18em] text-sidebar-foreground/45">Módulo activo</p>
                    <div className="rounded-lg border border-sidebar-border bg-sidebar-accent/70 p-3">
                    <div className="flex items-center gap-2">
                      <Calculator className="size-4 text-primary" />
                      <span className="text-sm font-medium">Utilidades</span>
                      </div>
                      <p className="mt-1 pl-6 text-xs text-sidebar-foreground/55">Cálculo y distribución</p>
                      </div>
                      </div>

                      <nav className="flex flex-col gap-1" aria-label="Navegación principal">{navItems.map(({ label, icon: Icon, count }) => <button key={label} onClick={() => go(label)} className={`flex items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm transition-colors ${view === label ? 'bg-primary text-primary-foreground shadow-sm' : 'text-sidebar-foreground/70 hover:bg-sidebar-accent hover:text-sidebar-foreground'}`}>
                        <Icon className="size-4" />
                        <span className="flex-1">{label}</span>{count && <span className={`rounded px-1.5 py-0.5 text-[10px] ${view === label ? 'bg-primary-foreground/15' : 'bg-sidebar-accent'}`}>{count}</span>}</button>)}</nav>
                        </div>
                        <div className="border-t border-sidebar-border p-3">
                          <div className="flex items-center gap-3 rounded-lg p-2">
                            <Avatar className="size-8">
                              <AvatarFallback className="bg-primary/10 text-xs text-primary">
                                {(user.nombre?.[0] ?? '') + (user.apellido?.[0] ?? '')}
                              </AvatarFallback>
                            </Avatar>
                              <div className="min-w-0 flex-1">
                                <p className="truncate text-xs font-medium">{user.nombre} {user.apellido}</p>
                                <p className="truncate text-[11px] text-sidebar-foreground/55 capitalize">{user.rol}</p>
                                </div>
                                <Button variant="ghost" size="icon" className="size-7" onClick={logout} aria-label="Cerrar sesión" title="Cerrar sesión">
                                  <LogOut className="size-3.5" />
                                </Button>
                                </div>
                                </div>
                                </aside>{mobileNav && <button aria-label="Cerrar menú" className="fixed inset-0 z-30 bg-foreground/20 lg:hidden" onClick={() => setMobileNav(false)} />}
<div className="lg:pl-64">
  <header className="sticky top-0 z-20 flex h-16 items-center gap-3 border-b border-border/70 bg-background/90 px-4 backdrop-blur sm:px-6">
  <Button variant="ghost" size="icon" className="lg:hidden" onClick={() => setMobileNav(true)}>
    <Menu />
    </Button>
    <div className="min-w-0 flex-1">
      <div className="flex items-center gap-2 text-xs text-muted-foreground">
        <span className="hidden sm:inline">Utilidades</span>
        <span className="hidden sm:inline">/</span>
        <span className="font-medium text-foreground">Ejercicio 2026</span>
        </div>
        <p className="mt-0.5 truncate text-xs text-muted-foreground">{company.name ? `${company.name} · RUC ${company.ruc}` : ''}</p>
        </div>
        <Tooltip>
          <TooltipTrigger render={<Button variant="ghost" size="icon" />}>
          <Bell className="size-4" />
          </TooltipTrigger>
          <TooltipContent>Notificaciones</TooltipContent>
          </Tooltip>
          <Button variant="outline" size="sm" className="hidden sm:flex" onClick={() => notify('Centro de ayuda abierto')}>
            <BookOpen className="mr-2 size-4" />Ayuda</Button>
            </header>
      <main className="mx-auto max-w-[1600px] p-4 sm:p-6 lg:p-8">
        <div className="mb-6 flex items-center justify-between gap-4">
          <div>
            <p className="text-xs font-medium text-muted-foreground">Periodo sin configurar</p>
            <p className="mt-1 text-sm font-medium">Cierre anual de utilidades</p>
            </div>
            </div>
      {view === 'Dashboard' && <DashboardView companies={companyClients} results={results} onNavigate={(next) => go(next)} />}      {view === 'Cálculo' && <CalculationView results={results} totals={totals} errors={validationErrors} onRun={() => setRunOpen(true)} onSelect={setSelected} onSummaryChange={setImportSummary} onReprocessed={handleReprocessed} />}
      {view === 'Empresas' && (
  <CompaniesView
    onOpen={() => go('Cálculo')}
    onNotify={notify}
    onShowAllModules={() => go('Configuración')}
  />
)}
      {view === 'Configuración' && <ConfigurationView onNotify={notify} />}
      {view === 'Trabajadores' && <WorkersView query={query} setQuery={setQuery} employees={filteredEmployees} onNotify={notify} />}
      {view === 'Resultados' && <ResultsView results={results} totals={totals} onSelect={setSelected} onNotify={notify} onConstancia={setConstancia} />}
      {view === 'Remanentes' && <RemaindersView results={results} totals={totals} />}
      {view === 'Reportes' && <ReportsView results={results} totals={totals} onNotify={notify} />}
      {view === 'Auditoría' && <AuditView onNotify={notify} />}
      </main></div>
      <Sheet open={!!selected} onOpenChange={(open) => !open && setSelected(null)}>
        <SheetContent className="w-full sm:max-w-lg">
          <SheetHeader>
            <SheetTitle>Detalle individual</SheetTitle>
            <SheetDescription>Resultado calculado para el ejercicio 2026.</SheetDescription>
            </SheetHeader>{selected && <div className="flex flex-col gap-6 p-5">
              <div className="flex items-center gap-3">
                <Avatar className="size-12">
                  <AvatarFallback className="bg-primary/10 text-primary">{selected.name.split(' ').map((n) => n[0]).slice(0,2).join('')}</AvatarFallback>
              </Avatar><div>
                <p className="font-semibold">{selected.name}</p>
                <p className="text-sm text-muted-foreground">{selected.role} · {selected.code}</p>
                </div></div>
                <div className="grid grid-cols-2 gap-3">{[['Días efectivos', number.format(selected.diasEfectivos)],['Remuneración computable', money.format(selected.remuneracionComputable)],['Utilidad por días', money.format(selected.distribucion.utilidadPorDias)],['Utilidad por rem.', money.format(selected.distribucion.utilidadPorRemuneraciones)]].map(([l,v]) => <div key={l} className="rounded-lg border border-border/70 p-3">
                <p className="text-xs text-muted-foreground">{l}</p>
                <p className="mt-1 font-semibold">{v}</p>
                </div>)}</div>
                <div className="rounded-xl bg-primary/5 p-4">
                <p className="text-sm text-muted-foreground">Utilidad bruta</p>
                <p className="mt-1 text-3xl font-semibold text-primary">{money.format(selected.distribucion.utilidadBruta)}</p>
                <div className="mt-3 flex justify-between text-xs">
                  <span>Preliminar</span>
                  <span>{money.format(selected.preliminary)}</span></div>
                  <div className="mt-1 flex justify-between text-xs">
                    <span>Quinta categoría</span>
                    <span>{money.format(selected.distribucion.quintaCategoria)}</span></div>
                  <div className="mt-1 flex justify-between text-xs">
                    <span>Tope ({parameters.capMonths} remuneraciones)</span>
                    <span>{money.format(selected.cap)}</span></div>
                  <div className="mt-3 flex justify-between border-t border-primary/20 pt-3 text-sm font-semibold">
                    <span>Utilidad neta</span>
                    <span>{money.format(selected.distribucion.utilidadNeta)}</span></div>
                    </div>
                    <Button onClick={() => { setConstancia(selected); setSelected(null) }}>
                      <FileText className="mr-2 size-4" />Ver boleta individual</Button>
                    </div>}</SheetContent></Sheet>
      <Sheet open={!!constancia} onOpenChange={(open) => !open && setConstancia(null)}>
        <SheetContent className="w-full overflow-y-auto sm:max-w-2xl">
          <SheetHeader>
            <SheetTitle>Boleta / Constancia individual</SheetTitle>
            <SheetDescription>Detalle del cálculo de utilidades para el trabajador.</SheetDescription>
          </SheetHeader>
          <div className="p-3">
            {constancia && <BoletaConstancia result={constancia} onNotify={notify} />}
          </div>
        </SheetContent>
      </Sheet>
      <Dialog open={runOpen} onOpenChange={setRunOpen}>
        <DialogContent><DialogHeader>
          <DialogTitle>Ejecutar cálculo de utilidades</DialogTitle>
          <DialogDescription>Revisa el resumen de la información extraída antes de generar el cálculo.</DialogDescription></DialogHeader>
          <div className="grid gap-3 py-2 sm:grid-cols-2">
            <div className="flex flex-col rounded-2xl border border-primary/20 bg-gradient-to-b from-primary/5 to-transparent p-5">
              <div className="flex items-center justify-between">
                <div className="flex size-10 items-center justify-center rounded-xl bg-primary/10 text-primary">
                  <Wallet className="size-5" />
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-[10px] font-semibold uppercase tracking-wider text-primary">Remuneraciones</span>
              </div>
              <p className="mt-5 text-xs text-muted-foreground">Remuneración anual computable</p>
              <p className="mt-1 text-2xl font-bold tracking-tight text-primary tabular-nums">{money.format(importSummary.totalRemuneracionAnual)}</p>
            </div>
            <div className="flex flex-col rounded-2xl border border-primary/20 bg-gradient-to-b from-primary/5 to-transparent p-5">
              <div className="flex items-center justify-between">
                <div className="flex size-10 items-center justify-center rounded-xl bg-primary/10 text-primary">
                  <CalendarDays className="size-5" />
                </div>
                <span className="rounded-full bg-primary/10 px-2.5 py-1 text-[10px] font-semibold uppercase tracking-wider text-primary">Días laborados</span>
              </div>
              <p className="mt-5 text-xs text-muted-foreground">Días registrados en el ejercicio</p>
              <p className="mt-1 text-2xl font-bold tracking-tight text-primary tabular-nums">{number.format(importSummary.totalDiasRegistrados)}</p>
              <div className="mt-4 flex items-center justify-between rounded-lg bg-background/60 px-3 py-2 text-sm">
                <span className="text-muted-foreground">Incidencias</span>
                <span className="font-semibold">{number.format(importSummary.totalIncidencias)}</span>
              </div>
            </div>
          </div>
            <div className="flex items-center gap-2 rounded-lg bg-amber-500/10 px-3 py-2.5 text-xs text-amber-500">
              <Info className="size-4 shrink-0" />El cálculo se registrará como una nueva versión al confirmar.
            </div>
            <DialogFooter>
              <Button variant="outline" onClick={() => setRunOpen(false)}>Cancelar</Button>
              <Button onClick={() => { setRunOpen(false); notify('Cálculo ejecutado correctamente') }}>
                <Play className="mr-2 size-4" />Confirmar ejecución</Button>
                </DialogFooter>
                </DialogContent>
                </Dialog>
      {toast && <div role="status" className="fixed bottom-5 right-5 z-50 flex items-center gap-3 rounded-lg border border-border bg-card px-4 py-3 text-sm shadow-lg">
        <Check className="size-4 text-emerald-400" />{toast}</div>}
    </div></TooltipProvider>
}

