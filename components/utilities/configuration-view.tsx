'use client'

import { Save, ShieldCheck } from 'lucide-react'

import { Button } from '@/components/ui/button'
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from '@/components/ui/card'
import { Input } from '@/components/ui/input'
import { company } from '@/lib/utilities-data'

function StatusBadge({
  type,
}: {
  type: 'configurable' | 'confirmado' | 'por-validar'
}) {
  const styles = {
    configurable:
      'border-blue-500/40 bg-blue-500/10 text-blue-400',
    confirmado:
      'border-emerald-500/40 bg-emerald-500/10 text-emerald-400',
    'por-validar':
      'border-amber-500/40 bg-amber-500/10 text-amber-400',
  }

  const labels = {
    configurable: 'CONFIGURABLE',
    confirmado: 'CONFIRMADO',
    'por-validar': 'POR VALIDAR',
  }

  return (
    <span
      className={`rounded-full border px-2 py-0.5 text-[9px] font-semibold uppercase tracking-wide ${styles[type]}`}
    >
      {labels[type]}
    </span>
  )
}

function SectionCard({
  title,
  description,
  children,
}: {
  title: string
  description: string
  children: React.ReactNode
}) {
  return (
    <Card className="overflow-hidden border-border/70 bg-card/80 shadow-none">
      <CardHeader className="border-b border-border/70 px-4 py-3.5 sm:px-5">
        <CardTitle className="text-sm font-semibold">
          {title}
        </CardTitle>

        <CardDescription className="text-[11px]">
          {description}
        </CardDescription>
      </CardHeader>

      <CardContent className="px-4 py-4 sm:px-5 sm:py-5">
        {children}
      </CardContent>
    </Card>
  )
}

function ConfigField({
  label,
  value,
  prefix,
  suffix,
  status,
  help,
}: {
  label: string
  value: string
  prefix?: string
  suffix?: string
  status?: 'configurable' | 'confirmado' | 'por-validar'
  help?: string
}) {
  return (
    <div className="space-y-2.5">
      <div className="flex items-center justify-between gap-3">
        <label className="text-xs font-medium text-foreground">
          {label}
        </label>

        {status && <StatusBadge type={status} />}
      </div>

      <div className="relative">
        {prefix && (
          <span className="pointer-events-none absolute left-3 top-1/2 z-10 -translate-y-1/2 text-xs text-muted-foreground">
            {prefix}
          </span>
        )}

        <Input
          type="number"
          defaultValue={value}
          className={`h-9 bg-background/40 text-sm ${
            prefix ? 'pl-9' : ''
          } ${suffix ? 'pr-12' : ''}`}
        />

        {suffix && (
          <span className="pointer-events-none absolute right-3 top-1/2 z-10 -translate-y-1/2 text-xs text-muted-foreground">
            {suffix}
          </span>
        )}
      </div>

      {help && (
        <p className="text-[10px] leading-4 text-muted-foreground">
          {help}
        </p>
      )}
    </div>
  )
}

export function ConfigurationView({
  onNotify,
}: {
  onNotify: (m: string) => void
}) {
  return (
    <div className="flex flex-col gap-6">
      {/* ENCABEZADO */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div>
          <p className="mb-1 text-[10px] font-semibold uppercase tracking-[0.18em] text-primary">
            Parámetros legales
          </p>

          <h1 className="text-2xl font-semibold tracking-tight text-foreground sm:text-3xl">
            Configuración
          </h1>

          <p className="mt-2 max-w-2xl text-xs leading-5 text-muted-foreground">
            Define la base legal y las reglas que utilizará el motor de
            cálculo.
          </p>
        </div>

        <Button
          onClick={() => onNotify('Configuración guardada')}
          className="h-9 gap-2"
        >
          <Save className="size-4" />
          Guardar cambios
        </Button>
      </div>

      {/* INFORMACIÓN DE EMPRESA */}
      <Card className="border-border/70 bg-card/80 shadow-none">
        <CardHeader className="border-b border-border/70 px-4 py-3.5 sm:px-5">
          <CardTitle className="text-sm font-semibold">
            Datos de la empresa
          </CardTitle>

          <CardDescription className="text-[11px]">
            Información tributaria del ejercicio.
          </CardDescription>
        </CardHeader>

        <CardContent className="grid gap-5 px-4 py-4 sm:grid-cols-2 sm:px-5">
          <div>
            <p className="text-[10px] text-muted-foreground">
              Razón social
            </p>

            <p className="mt-1 text-sm font-medium">
              {company.name}
            </p>
          </div>

          <div>
            <p className="text-[10px] text-muted-foreground">
              Actividad económica
            </p>

            <p className="mt-1 text-sm font-medium">
              {company.activity}
            </p>
          </div>

          <div className="sm:col-span-2">
            <ConfigField
              label="Renta neta anual"
              value=""
              prefix="S/"
              status="por-validar"
              help="Ingresa la renta neta anual correspondiente al ejercicio."
            />
          </div>
        </CardContent>
      </Card>

      {/* GRID PRINCIPAL */}
      <div className="grid gap-5 lg:grid-cols-2">
        {/* DISTRIBUCIÓN */}
        <SectionCard
          title="Distribución"
          description="Monto y porcentaje a repartir"
        >
          <div className="space-y-5">
            <ConfigField
              label="Porcentaje de distribución"
              value="10"
              suffix="%"
              status="configurable"
              help="Porcentaje de la renta neta anual destinado a los trabajadores (5% a 10% según actividad)."
            />

            <ConfigField
              label="Monto total a distribuir"
              value="4860000"
              prefix="S/"
              status="por-validar"
              help="Se obtiene de la renta neta imponible. Verifica contra la declaración anual."
            />
          </div>
        </SectionCard>

        {/* FACTORES DE REPARTO */}
        <SectionCard
          title="Factores de reparto"
          description="Ponderación entre días y remuneraciones"
        >
          <div className="space-y-5">
            <div className="space-y-2.5">
              <div className="flex items-center justify-between gap-3">
                <label className="text-xs font-medium text-foreground">
                  Reparto por días trabajados
                </label>

                <StatusBadge type="configurable" />
              </div>

              <div className="flex items-center gap-3">
                <Input
                  type="number"
                  defaultValue="50"
                  className="h-9 w-24 bg-background/40 text-sm"
                />

                <span className="text-[11px] text-muted-foreground">
                  % por días · 50% por remuneraciones
                </span>
              </div>

              <div className="flex h-1.5 overflow-hidden rounded-full bg-muted">
                <div
                  className="bg-primary"
                  style={{ width: '50%' }}
                />

                <div
                  className="bg-chart-2"
                  style={{ width: '50%' }}
                />
              </div>

              <p className="text-[10px] leading-4 text-muted-foreground">
                La ley establece 50% por días trabajados y 50% en
                proporción a las remuneraciones.
              </p>
            </div>
          </div>
        </SectionCard>

        {/* QUINTA CATEGORÍA */}
        <SectionCard
          title="Quinta categoría"
          description="Parámetros del impuesto a la renta"
        >
          <div className="space-y-5">
            <ConfigField
              label="Valor de la UIT"
              value="5350"
              prefix="S/"
              status="confirmado"
              help="UIT vigente para el ejercicio 2026."
            />

            <ConfigField
              label="Deducción (UIT)"
              value="7"
              suffix="UIT"
              status="confirmado"
              help="Deducción de 7 UIT sobre las rentas de cuarta y quinta categoría."
            />

            <p className="text-xs text-muted-foreground">
              Deducción equivalente a{' '}
              <span className="font-semibold text-foreground">
                S/ 37,450.00
              </span>
            </p>
          </div>
        </SectionCard>

        {/* ESCALA PROGRESIVA */}
        <SectionCard
          title="Escala progresiva"
          description="Tramos del impuesto de quinta categoría"
        >
          <div className="divide-y divide-border">
            <div className="flex items-center justify-between py-2.5 text-xs">
              <span className="text-muted-foreground">
                De 0 a 5 UIT
              </span>

              <span className="font-semibold">8%</span>
            </div>

            <div className="flex items-center justify-between py-2.5 text-xs">
              <span className="text-muted-foreground">
                De 5 a 20 UIT
              </span>

              <span className="font-semibold">14%</span>
            </div>

            <div className="flex items-center justify-between py-2.5 text-xs">
              <span className="text-muted-foreground">
                De 20 a 35 UIT
              </span>

              <span className="font-semibold">17%</span>
            </div>

            <div className="flex items-center justify-between py-2.5 text-xs">
              <span className="text-muted-foreground">
                De 35 a 45 UIT
              </span>

              <span className="font-semibold">20%</span>
            </div>

            <div className="flex items-center justify-between py-2.5 text-xs">
              <span className="text-muted-foreground">
                Más de 45 UIT
              </span>

              <span className="font-semibold">30%</span>
            </div>
          </div>
        </SectionCard>
      </div>
    </div>
  )
}