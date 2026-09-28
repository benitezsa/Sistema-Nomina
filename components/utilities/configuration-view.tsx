'use client'

import { useCallback, useEffect, useRef, useState } from 'react'
import { Save } from 'lucide-react'

import { Button } from '@/components/ui/button'
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from '@/components/ui/card'
import { Input } from '@/components/ui/input'
import { useSettings } from '@/lib/settings-context'

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
  onChange,
  disabled,
}: {
  label: string
  value: string
  prefix?: string
  suffix?: string
  status?: 'configurable' | 'confirmado' | 'por-validar'
  help?: string
  onChange?: (value: string) => void
  disabled?: boolean
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
          type="text"
          inputMode="decimal"
          value={value}
          onChange={(event) => onChange?.(event.target.value)}
          disabled={disabled}
          className={`h-9 bg-background/40 text-sm ${
            prefix ? 'pl-9' : ''
          } ${suffix ? 'pr-12' : ''} ${disabled ? 'opacity-70' : ''}`}
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
  const ANIO_EJERCICIO = 2026

  const { company, parameters, updateParameters } = useSettings()

  const [income, setIncome] = useState(
    parameters.income > 0 ? String(parameters.income) : '',
  )

  const editado = useRef(false)

  useEffect(() => {
    setIncome(parameters.income > 0 ? String(parameters.income) : '')
  }, [parameters.income])

  useEffect(() => {
    const empresaId = Number(company.id)

    if (!empresaId) return

    let activo = true

    fetch(`/api/ejercicios?empresa_id=${empresaId}&anio=${ANIO_EJERCICIO}`)
      .then((response) => response.json())
      .then((data) => {
        if (!activo) return

        if (!data?.existe || !data.ejercicio) {
          setIncome('')
          return
        }

        const renta = Number(data.ejercicio.renta_neta)
        const porcentaje = Number(data.ejercicio.porcentaje_distribucion)

        setIncome(renta > 0 ? String(renta) : '')
        updateParameters({ income: renta, legalPercent: porcentaje })
      })
      .catch(() => {})

    return () => {
      activo = false
    }
  }, [company.id, updateParameters])

  const legalPercent = String(parameters.legalPercent)

  const sanitizeMoney = useCallback(
    (value: string) =>
      String(value ?? '')
        .replace(/[^\d.]/g, '')
        .replace(/(\..*)\./g, '$1'),
    [],
  )

  const montoTotal =
    Number(sanitizeMoney(income)) > 0 && Number(legalPercent) > 0
      ? Number(sanitizeMoney(income)) * (Number(legalPercent) / 100)
      : 0

  const formatMoney = (value: number) =>
    value.toLocaleString('es-PE', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })

  const montoDias = montoTotal > 0 ? montoTotal / 2 : 0
  const montoRemuneraciones = montoTotal > 0 ? montoTotal / 2 : 0

  const guardarValores = useCallback(async (): Promise<boolean> => {
    const empresaId = Number(company.id)

    if (!empresaId) return false

    const renta = Number(sanitizeMoney(income))
    const porcentaje = Number(sanitizeMoney(legalPercent))

    if (!(renta > 0) || !(porcentaje > 0)) return false

    try {
      const response = await fetch('/api/ejercicios', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          empresa_id: empresaId,
          anio: ANIO_EJERCICIO,
          renta_neta: renta,
          porcentaje_distribucion: porcentaje,
        }),
      })

      const data = await response.json()

      if (!response.ok || !data?.success) return false

      updateParameters({ income: renta, legalPercent: porcentaje })
      return true
    } catch {
      return false
    }
  }, [company.id, income, legalPercent, sanitizeMoney, updateParameters])

  useEffect(() => {
    if (!editado.current) return

    const timer = window.setTimeout(() => {
      void guardarValores()
    }, 800)

    return () => window.clearTimeout(timer)
  }, [income, legalPercent, guardarValores])

  const guardar = async () => {
    const empresaId = Number(company.id)

    if (!empresaId) {
      onNotify('Selecciona una empresa antes de guardar')
      return
    }

    const renta = Number(sanitizeMoney(income))
    const porcentaje = Number(sanitizeMoney(legalPercent))

    if (!(renta > 0) || !(porcentaje > 0)) {
      onNotify('Ingresa la renta neta y el porcentaje de distribución')
      return
    }

    const ok = await guardarValores()
    onNotify(ok ? 'Configuración guardada' : 'Error al guardar la configuración')
  }

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
            cálculo. Los valores se adaptan a la empresa seleccionada.
          </p>
        </div>

        <Button
          onClick={guardar}
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
              {company.name || 'Sin empresa seleccionada'}
            </p>
          </div>

          <div>
            <p className="text-[10px] text-muted-foreground">
              Actividad económica
            </p>

            <p className="mt-1 text-sm font-medium">
              {company.activity || '—'}
            </p>
          </div>

          <div className="sm:col-span-2">
            <ConfigField
              label="Renta neta anual"
              value={income}
              prefix="S/"
              status="por-validar"
              onChange={(value) => {
                editado.current = true
                setIncome(sanitizeMoney(value))
              }}
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
              value={legalPercent}
              suffix="%"
              status="confirmado"
              onChange={(value) => {
                editado.current = true
                updateParameters({
                  legalPercent: Number(sanitizeMoney(value)),
                })
              }}
              help={`Porcentaje de la renta neta anual destinado a los trabajadores (5% a 10% según actividad). Se define por el rubro de la empresa ${company.activity ? `"${company.activity}"` : 'seleccionada'}.`}
            />

            <ConfigField
              label="Monto total a distribuir"
              value={montoTotal > 0 ? formatMoney(montoTotal) : ''}
              prefix="S/"
              status="por-validar"
              disabled
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

              <ConfigField
                label="Monto destinado a días (50%)"
                value={montoDias > 0 ? formatMoney(montoDias) : ''}
                prefix="S/"
                status="por-validar"
                disabled
              />

              <ConfigField
                label="Monto destinado a remuneraciones (50%)"
                value={montoRemuneraciones > 0 ? formatMoney(montoRemuneraciones) : ''}
                prefix="S/"
                status="por-validar"
                disabled
              />

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
                La ley establece 50% del monto total a distribuir por días
                trabajados y 50% en proporción a las remuneraciones.
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
              disabled
              help="UIT vigente para el ejercicio 2026."
            />

            <ConfigField
              label="Deducción (UIT)"
              value="7"
              suffix="UIT"
              status="confirmado"
              disabled
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