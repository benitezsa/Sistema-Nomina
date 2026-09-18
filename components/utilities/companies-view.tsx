'use client'

import { useMemo, useState } from 'react'
import { ArrowRight, Building2, Check, Mail, MapPin, Phone, Plus, Search, Users } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '@/components/ui/card'
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog'
import { Input } from '@/components/ui/input'
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select'
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from '@/components/ui/sheet'
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table'
import { companyClients } from '@/lib/utilities-data'
import { useSettings } from '@/lib/settings-context'
import type { CompanyClient } from './types'

const statusStyle: Record<string, string> = {
  Activo: 'bg-emerald-500/10 text-emerald-400 border-emerald-500/30',
  Pendiente: 'bg-amber-500/10 text-amber-400 border-amber-500/30',
  Inactivo: 'bg-slate-500/10 text-slate-300 border-slate-500/30',
}

const money = new Intl.NumberFormat('es-PE')

const rubroOptions = [
  { value: 'Pesquera', label: 'Empresas pesqueras', porcentaje: '10%' },
  { value: 'Telecomunicaciones', label: 'Empresas de telecomunicaciones', porcentaje: '10%' },
  { value: 'Industrial', label: 'Empresas industriales', porcentaje: '10%' },
  { value: 'Minera', label: 'Empresas mineras', porcentaje: '8%' },
  { value: 'Comercio', label: 'Comercio y restaurantes', porcentaje: '8%' },
  { value: 'Otras', label: 'Otras actividades', porcentaje: '5%' },
]

function Status({ value }: { value: string }) {
  return (
    <Badge
      variant="outline"
      className={`font-medium ${statusStyle[value] || ''}`}
    >
      {value}
    </Badge>
  )
}

function Kpi({
  label,
  value,
  hint,
  icon: Icon,
}: {
  label: string
  value: string
  hint: string
  icon: typeof Building2
}) {
  return (
    <Card className="border-border/70 shadow-sm">
      <CardContent className="flex items-start justify-between p-5">
        <div>
          <p className="text-xs font-medium text-muted-foreground">{label}</p>
          <p className="mt-2 text-xl font-semibold tracking-tight">{value}</p>
          <p className="mt-1 text-xs text-teal-400">{hint}</p>
        </div>

        <div className="rounded-lg bg-primary/8 p-2.5 text-primary">
          <Icon className="size-4" />
        </div>
      </CardContent>
    </Card>
  )
}

export function CompaniesView({
  onOpen,
  onNotify,
  onShowAllModules,
}: {
  onOpen: () => void
  onNotify: (message: string) => void
  onShowAllModules: () => void
}) {
  const [query, setQuery] = useState('')
  const [status, setStatus] = useState('Todos')
  const [selected, setSelected] = useState<CompanyClient | null>(null)
  const [dialogOpen, setDialogOpen] = useState(false)
  const [rubro, setRubro] = useState<string>('')
  const [companies, setCompanies] = useState<CompanyClient[]>(companyClients)
  const [form, setForm] = useState({ legalName: '', ruc: '', tradeName: '', contact: '' })

  const { setCompany, updateParameters } = useSettings()

  const porcentajeDeRubro = (activity: string) =>
    rubroOptions.find((option) => option.label === activity)?.porcentaje ?? null

  const activarEmpresa = (item: CompanyClient) => {
    const porcentaje = porcentajeDeRubro(item.activity)
    const number = porcentaje ? Number.parseFloat(porcentaje) : null

    setCompany({
      id: item.id,
      name: item.legalName,
      ruc: item.ruc,
      activity: item.activity,
      address: item.address,
    })

    if (number) {
      updateParameters({ legalPercent: number })
    }

    return number
  }

  const porcentajeAutomatico = useMemo(
    () => rubroOptions.find((option) => option.value === rubro)?.porcentaje ?? '',
    [rubro],
  )

  const filtered = useMemo(
    () =>
      companies.filter(
        (item) =>
          `${item.legalName} ${item.tradeName} ${item.ruc} ${item.contact}`
            .toLowerCase()
            .includes(query.toLowerCase()) &&
          (status === 'Todos' || item.status === status),
      ),
    [query, status, companies],
  )

  const active = companies.filter(
    (item) => item.status === 'Activo',
  ).length

  const employees = companies.reduce(
    (sum, item) => sum + item.employees,
    0,
  )

  const registerCompany = () => {
    if (!form.legalName.trim() || !form.ruc.trim()) {
      onNotify('Razón social y RUC son obligatorios')
      return
    }

    const next: CompanyClient = {
      id: `emp-${Date.now()}`,
      legalName: form.legalName.trim(),
      tradeName: form.tradeName.trim() || form.legalName.trim(),
      ruc: form.ruc.trim(),
      activity:
        rubroOptions.find((option) => option.value === rubro)?.label || rubro || 'Sin rubro',
      contact: form.contact.trim(),
      role: 'Contacto principal',
      phone: '',
      email: '',
      address: '',
      status: 'Pendiente',
      employees: 0,
      lastExercise: 0,
    }

    setCompanies((prev) => [next, ...prev])
    setForm({ legalName: '', ruc: '', tradeName: '', contact: '' })
    setRubro('')
    setDialogOpen(false)
    onNotify('Empresa registrada en el Directorio de empresas')
  }

  return (
    <div className="flex flex-col gap-7">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="mb-1 text-xs font-semibold uppercase tracking-[0.18em] text-primary">
            Gestión de clientes
          </p>

          <h1 className="text-2xl font-semibold tracking-tight sm:text-3xl">
            Empresas
          </h1>

          <p className="mt-2 max-w-2xl text-sm leading-6 text-muted-foreground">
            Administra las empresas afiliadas y consulta su información
            tributaria, contactos y estado operativo.
          </p>
        </div>

        <Button onClick={() => setDialogOpen(true)}>
          <Plus className="mr-2 size-4" />
          Nueva empresa
        </Button>
      </div>

      <div className="grid gap-4 sm:grid-cols-3">
        <Kpi
          label="Empresas registradas"
          value={money.format(companies.length)}
          hint="Clientes en la plataforma"
          icon={Building2}
        />

        <Kpi
          label="Empresas activas"
          value={money.format(active)}
          hint="Con operación vigente"
          icon={Check}
        />

        <Kpi
          label="Trabajadores asociados"
          value={money.format(employees)}
          hint="En todas las empresas"
          icon={Users}
        />
      </div>

      <Card>
        <CardHeader>
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div>
              <CardTitle>Directorio de empresas</CardTitle>

              <CardDescription>
                Busca por razón social, nombre comercial, RUC o contacto.
              </CardDescription>
            </div>

            <div className="flex flex-col gap-2 sm:flex-row">
              <div className="relative">
                <Search className="absolute left-3 top-1/2 size-4 -translate-y-1/2 text-muted-foreground" />

                <Input
                  value={query}
                  onChange={(event) => setQuery(event.target.value)}
                  placeholder="Buscar empresa..."
                  className="pl-9 sm:w-72"
                />
              </div>

              <Select
  value={status}
  onValueChange={(value) => setStatus(value ?? 'Todos')}
>
                <SelectTrigger className="sm:w-36">
                  <SelectValue />
                </SelectTrigger>

                <SelectContent>
                  <SelectItem value="Todos">Todos los estados</SelectItem>
                  <SelectItem value="Activo">Activo</SelectItem>
                  <SelectItem value="Pendiente">Pendiente</SelectItem>
                  <SelectItem value="Inactivo">Inactivo</SelectItem>
                </SelectContent>
              </Select>
            </div>
          </div>
        </CardHeader>

        <CardContent className="p-0">
          <div className="overflow-x-auto">
            {filtered.length ? (
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Empresa</TableHead>
                    <TableHead>RUC</TableHead>
                    <TableHead>Contacto</TableHead>
                    <TableHead>Trabajadores</TableHead>
                    <TableHead>Estado</TableHead>
                    <TableHead className="text-right">Acción</TableHead>
                  </TableRow>
                </TableHeader>

                <TableBody>
                  {filtered.map((item) => (
                    <TableRow key={item.id}>
                      <TableCell>
                        <p className="font-medium">{item.legalName}</p>

                        <p className="text-xs text-muted-foreground">
                          {item.tradeName} · {item.activity}
                        </p>
                      </TableCell>

                      <TableCell className="font-mono text-xs">
                        {item.ruc}
                      </TableCell>

                      <TableCell>
                        <p className="font-medium">{item.contact}</p>

                        <p className="text-xs text-muted-foreground">
                          {item.email}
                        </p>
                      </TableCell>

                      <TableCell>{item.employees}</TableCell>

                      <TableCell>
                        <Status value={item.status} />
                      </TableCell>

                      <TableCell className="text-right">
                        <div className="flex justify-end gap-2">
                          {/* BOTÓN ORIGINAL: NO SE MODIFICÓ */}
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => setSelected(item)}
                          >
                            Ver detalle
                            <ArrowRight className="ml-2 size-4" />
                          </Button>

                          {/* NUEVO BOTÓN */}
                          <Button
  variant="outline"
  size="sm"
 onClick={() => {
  const pct = activarEmpresa(item)
  localStorage.setItem('utilities-modules-unlocked', 'true')
  onShowAllModules()
  onNotify(
    pct
      ? `Módulos habilitados para ${item.legalName} · ${pct} de distribución aplicado`
      : `Módulos habilitados para ${item.legalName}`
  )
}}
>
  Detalles
</Button>
                        </div>
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            ) : (
              <div className="flex flex-col items-center justify-center gap-3 px-6 py-16 text-center">
                <div className="rounded-full bg-muted p-3">
                  <Search className="size-5 text-muted-foreground" />
                </div>

                <p className="font-medium">No encontramos empresas</p>

                <p className="max-w-sm text-sm text-muted-foreground">
                  Prueba con otro término de búsqueda o cambia el filtro de
                  estado.
                </p>

                <Button
                  variant="outline"
                  onClick={() => {
                    setQuery('')
                    setStatus('Todos')
                  }}
                >
                  Limpiar filtros
                </Button>
              </div>
            )}
          </div>
        </CardContent>
      </Card>

      <Dialog open={dialogOpen} onOpenChange={setDialogOpen}>
        <DialogContent>
          <DialogHeader>
            <DialogTitle>Nueva empresa</DialogTitle>

            <DialogDescription>
              Registra una empresa para habilitar sus ejercicios y
              trabajadores.
            </DialogDescription>
          </DialogHeader>

          <div className="grid gap-4 sm:grid-cols-2">
            <div className="sm:col-span-2">
              <label
                className="text-sm font-medium"
                htmlFor="legalName"
              >
                Razón social
              </label>

              <Input
                id="legalName"
                className="mt-2"
                placeholder="Ingrese la razón social"
                value={form.legalName}
                onChange={(event) => setForm((prev) => ({ ...prev, legalName: event.target.value }))}
              />
            </div>

            <div>
              <label className="text-sm font-medium" htmlFor="ruc">
                RUC
              </label>

              <Input
                id="ruc"
                className="mt-2"
                placeholder="Ingrese el RUC"
                value={form.ruc}
                onChange={(event) => setForm((prev) => ({ ...prev, ruc: event.target.value }))}
              />
            </div>

            <div>
              <label
                className="text-sm font-medium"
                htmlFor="tradeName"
              >
                Nombre comercial
              </label>

              <Input
                id="tradeName"
                className="mt-2"
                placeholder="Nombre comercial"
                value={form.tradeName}
                onChange={(event) => setForm((prev) => ({ ...prev, tradeName: event.target.value }))}
              />
            </div>

            <div className="sm:col-span-2">
              <label
                className="text-sm font-medium"
                htmlFor="contact"
              >
                Contacto principal
              </label>

              <Input
                id="contact"
                className="mt-2"
                placeholder="Nombre y cargo"
                value={form.contact}
                onChange={(event) => setForm((prev) => ({ ...prev, contact: event.target.value }))}
              />
            </div>

            <div className="sm:col-span-2">
              <label className="text-sm font-medium" htmlFor="rubro">
                Rubro al que pertenece
              </label>

              <Select value={rubro} onValueChange={(value) => setRubro(value ?? '')}>
                <SelectTrigger id="rubro" className="mt-2">
                  <SelectValue placeholder="Seleccione el rubro" />
                </SelectTrigger>

                <SelectContent>
                  {rubroOptions.map((option) => (
                    <SelectItem key={option.value} value={option.value}>
                      {option.label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            </div>

            <div className="sm:col-span-2">
              <label
                className="text-sm font-medium"
                htmlFor="porcentaje"
              >
                Porcentaje de participación
              </label>

              <Input
                id="porcentaje"
                className="mt-2"
                value={porcentajeAutomatico}
                placeholder="Se completa automáticamente"
                disabled
              />
            </div>
          </div>

          <DialogFooter>
            <Button
              variant="outline"
              onClick={() => setDialogOpen(false)}
            >
              Cancelar
            </Button>

            <Button
              onClick={registerCompany}
            >
              Guardar empresa
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <Sheet
        open={!!selected}
        onOpenChange={(open) => !open && setSelected(null)}
      >
        <SheetContent className="w-full sm:max-w-lg">
          <SheetHeader>
            <SheetTitle>Detalle de empresa</SheetTitle>

            <SheetDescription>
              Información general y configuración del cliente.
            </SheetDescription>
          </SheetHeader>

          {selected && (
            <div className="flex flex-col gap-6 p-5">
              <div className="flex items-start gap-3">
                <div className="flex size-12 items-center justify-center rounded-xl bg-primary/10 text-primary">
                  <Building2 className="size-5" />
                </div>

                <div className="min-w-0">
                  <p className="font-semibold">{selected.legalName}</p>

                  <p className="text-sm text-muted-foreground">
                    {selected.tradeName} · RUC {selected.ruc}
                  </p>

                  <div className="mt-2">
                    <Status value={selected.status} />
                  </div>
                </div>
              </div>

              <div className="flex flex-col gap-3 rounded-xl border border-border/70 p-4 text-sm">
                <div className="flex gap-3">
                  <MapPin className="mt-0.5 size-4 text-muted-foreground" />
                  <span>{selected.address}</span>
                </div>

                <div className="flex gap-3">
                  <Phone className="mt-0.5 size-4 text-muted-foreground" />
                  <span>{selected.phone}</span>
                </div>

                <div className="flex gap-3">
                  <Mail className="mt-0.5 size-4 text-muted-foreground" />
                  <span className="break-all">{selected.email}</span>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div className="rounded-lg border border-border/70 p-3">
                  <p className="text-xs text-muted-foreground">
                    Trabajadores
                  </p>

                  <p className="mt-1 text-xl font-semibold">
                    {selected.employees}
                  </p>
                </div>

                <div className="rounded-lg border border-border/70 p-3">
                  <p className="text-xs text-muted-foreground">
                    Último ejercicio
                  </p>

                  <p className="mt-1 text-xl font-semibold">
                    {selected.lastExercise}
                  </p>
                </div>
              </div>

              <div className="rounded-xl bg-muted/50 p-4">
                <p className="text-sm font-medium">Próximamente</p>

                <p className="mt-1 text-sm text-muted-foreground">
                  Ejercicios, parámetros legales, nómina y reportes asociados
                  a esta empresa.
                </p>
              </div>

              <Button
                onClick={() => {
                  setSelected(null)
                  onOpen()
                }}
              >
                Abrir cálculo de utilidades
                <ArrowRight className="ml-2 size-4" />
              </Button>
            </div>
          )}
        </SheetContent>
      </Sheet>
    </div>
  )
}