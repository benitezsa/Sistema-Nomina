'use client'

import { ClipboardCheck } from 'lucide-react'

export function ValidationView() {
  return (
    <div className="rounded-xl border border-border/60 bg-card overflow-hidden">
      <div className="border-b border-border/60 px-5 py-4">
        <div className="flex items-center gap-3">
          <div className="flex size-10 items-center justify-center rounded-lg bg-primary/10 text-primary">
            <ClipboardCheck className="size-5" />
          </div>

          <div>
            <p className="text-sm font-semibold">
              Validación de trabajadores
            </p>

            <p className="mt-1 text-xs text-muted-foreground">
              Revisa la información cargada antes de continuar con el cálculo.
            </p>
          </div>
        </div>
      </div>

      <div className="p-5">
        <div className="rounded-lg border border-border/60 bg-muted/10 p-6 text-center">
          <p className="text-sm font-medium">
            Apartado de validación
          </p>

          <p className="mt-1 text-xs text-muted-foreground">
            Aquí mostraremos los trabajadores y el detalle de la información
            importada.
          </p>
        </div>
      </div>
    </div>
  )
}