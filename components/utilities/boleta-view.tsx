
'use client'

import { Download, Printer } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { money, number } from '@/lib/utilities-calculation'
import { useSettings } from '@/lib/settings-context'
import { buildBoletaHtml, exportBoleta, openPrintWindow } from '@/lib/report-export'
import type { EmployeeUtilityResult } from './types'

export function BoletaConstancia({
  result,
  onNotify,
}: {
  result: EmployeeUtilityResult
  onNotify: (m: string) => void
}) {
  const { company, parameters } = useSettings()
  const r = result

  const handlePrint = () => {
    openPrintWindow(`Constancia - ${r.name}`, buildBoletaHtml({ company, parameters }, r))
    onNotify('Vista de impresión abierta')
  }

  const handleDownload = async () => {
    try {
      await exportBoleta({ company, parameters }, r)
      onNotify('Constancia descargada (Excel)')
    } catch {
      onNotify('No se pudo descargar la constancia')
    }
  }

  const fila = (label: string, valor: string, extra?: React.ReactNode) => (
    <div className="flex items-center justify-between border-b border-border/60 py-2.5">
      <span className="text-xs text-muted-foreground">{label}</span>
      <span className="text-sm font-medium">{valor}{extra}</span>
    </div>
  )

  return (
    <div className="flex flex-col gap-0">
      {/* Encabezado de la constancia */}
      <div className="rounded-t-xl border border-border/60 bg-card px-6 pt-6">
        <div className="flex items-end justify-between border-b border-border pb-4">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.18em] text-primary">Constancia de utilidades</p>
            <h2 className="mt-1 text-xl font-semibold">Boleta individual de distribución</h2>
            <p className="mt-1 text-xs text-muted-foreground">Ejercicio 2026 · Utilidades de los trabajadores</p>
          </div>
          {company.name && (
            <div className="text-right">
              <p className="text-sm font-semibold">{company.name}</p>
              <p className="text-xs text-muted-foreground">RUC {company.ruc}</p>
            </div>
          )}
        </div>
      </div>

      <div className="border-x border-border/60 bg-card px-6">
        {/* Datos del trabajador */}
        <div className="grid gap-4 py-5 sm:grid-cols-3">
          <div className="sm:col-span-1">
            <p className="text-[10px] uppercase tracking-wide text-muted-foreground">Trabajador</p>
            <p className="mt-1 text-sm font-semibold">{r.name}</p>
            <p className="text-xs text-muted-foreground">{r.code} · {r.department}</p>
          </div>
          <div>
            <p className="text-[10px] uppercase tracking-wide text-muted-foreground">Cargo</p>
            <p className="mt-1 text-sm font-medium">{r.role}</p>
          </div>
          <div>
            <p className="text-[10px] uppercase tracking-wide text-muted-foreground">Jornada</p>
            <p className="mt-1 text-sm font-medium">{r.jornada ? `${r.jornada} días/semana` : '6 días/semana'}</p>
          </div>
        </div>

        {/* Datos base del cálculo */}
        <div className="rounded-xl border border-border/70 bg-muted/20 p-4">
          <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Datos base del cálculo</p>
          {fila('Días laborables', number.format(r.diasLaborables ?? 0))}
          {fila('Días no laborados', number.format(r.diasNoLaborados ?? 0))}
          {fila('Días efectivos', number.format(r.diasEfectivos), <span className="ml-1 text-primary">· factor { (r.daysFactor * 100).toFixed(2) }%</span>)}
          {fila('Remuneración computable', money.format(r.remuneracionComputable), <span className="ml-1 text-primary">· factor { (r.remunerationFactor * 100).toFixed(2) }%</span>)}
        </div>
      </div>

      {/* Desglose del cálculo */}
      <div className="border-x border-border/60 bg-card px-6">
        <div className="grid gap-0 sm:grid-cols-2">
          <div className="py-5 sm:pr-5">
            <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Distribución (50% / 50%)</p>
            {fila('Utilidad por días efectivos', money.format(r.distribucion.utilidadPorDias))}
            {fila('Utilidad por remuneraciones', money.format(r.distribucion.utilidadPorRemuneraciones))}
            {fila('Utilidad bruta', money.format(r.distribucion.utilidadBruta))}
            {r.capApplied && fila('Tope aplicado (18 remuneraciones)', money.format(r.cap))}
            {r.capApplied && fila('Excedente', money.format(r.remainder))}
          </div>
          <div className="border-t border-border/60 py-5 sm:border-l sm:border-t-0 sm:pl-5">
            <p className="mb-3 text-xs font-semibold uppercase tracking-wide text-muted-foreground">Quinta categoría</p>
            {fila('Remuneración anual', money.format(r.quintaCategoria?.remuneracionAnual ?? r.distribucion.utilidadBruta))}
            {fila('Deducción 7 UIT', `- ${money.format(r.quintaCategoria?.deduccion7UIT ?? 0)}`)}
            {fila('Base imponible', money.format(r.quintaCategoria?.baseImponible ?? 0))}
            {fila('Impuesto quinta categoría', money.format(r.distribucion.quintaCategoria))}
          </div>
        </div>
      </div>

      {/* Total neto */}
      <div className="rounded-b-xl border border-t-0 border-border/60 bg-primary/5 px-6 py-5">
        <div className="flex items-center justify-between">
          <div>
            <p className="text-xs font-medium text-muted-foreground">Total a pagar al trabajador</p>
            <p className="text-xs text-muted-foreground">Utilidad bruta menos quinta categoría</p>
          </div>
          <p className="text-2xl font-semibold text-primary">{money.format(r.distribucion.utilidadNeta)}</p>
        </div>
      </div>

      {/* Acciones */}
      <div className="mt-4 flex justify-end gap-2">
        <Button variant="outline" size="sm" onClick={handlePrint}>
          <Printer className="mr-2 size-4" />Imprimir
        </Button>
        <Button size="sm" onClick={handleDownload}>
          <Download className="mr-2 size-4" />Descargar
        </Button>
      </div>
    </div>
  )
}
