'use client'

import { useRef, useState } from 'react'
import ExcelJS from 'exceljs'

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

type MonthlyValues = {
  enero: number
  febrero: number
  marzo: number
  abril: number
  mayo: number
  junio: number
  julio: number
  agosto: number
  septiembre: number
  octubre: number
  noviembre: number
  diciembre: number
  total: number
}

type ImportedWorker = {
  dni: string
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string
  fechaInicio: ExcelJS.CellValue
  fechaCese: ExcelJS.CellValue
  remuneraciones: MonthlyValues
  diasTrabajados?: MonthlyValues
}

const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']   as const

type PendingDays = Map<string, MonthlyValues>

function ImportCard({
  kind,
  trabajadores,
  setTrabajadores,
}: {
  kind: 'remunerations' | 'worked-days'
  trabajadores: ImportedWorker[]
  setTrabajadores: React.Dispatch<
    React.SetStateAction<ImportedWorker[]>
  >
}) {
  const isRemunerations = kind === 'remunerations'

  const title = isRemunerations
    ? 'Remuneraciones'
    : 'Días laborados'

  const description = isRemunerations
    ? 'Importa la información mensual o consolidada de remuneraciones computables.'
    : 'Carga los días trabajados, ausencias y licencias del ejercicio.'

  const inputRef = useRef<HTMLInputElement>(null)
  const [loading, setLoading] = useState(false)

const [diasImportados, setDiasImportados] = useState(0)
const [diasPendientes, setDiasPendientes] = useState<
  Map<string, MonthlyValues>
>(new Map())

  const [importedFile, setImportedFile] = useState<string | null>(null)
const [trabajadorSeleccionado, setTrabajadorSeleccionado] = useState<number | null>(null)
const formatExcelDate = (value: ExcelJS.CellValue) => {
  if (!value) return '—'

  if (value instanceof Date) {
    const day = String(value.getUTCDate()).padStart(2, '0')
    const month = String(value.getUTCMonth() + 1).padStart(2, '0')
    const year = value.getUTCFullYear()

    return `${day}/${month}/${year}`
  }

  return String(value)
}

  const handleFile = async (file: File) => {
    setLoading(true)

    try {
      const workbook = new ExcelJS.Workbook()
      const buffer = await file.arrayBuffer()

      await workbook.xlsx.load(buffer)

      if (!isRemunerations) {
        const normalizeHeader = (value: ExcelJS.CellValue) => {
          const text = typeof value === 'object' && value !== null && 'text' in value
            ? String(value.text ?? '')
            : String(value ?? '')
          const normalized = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().replace(/\s+/g, ' ').toUpperCase()
          return normalized === 'SETIEMBRE' ? 'SEPTIEMBRE' : normalized
        }
        const normalizeDni = (value: ExcelJS.CellValue) => {
          const dni = String(value ?? '').trim().replace(/\.0$/, '')
          return /^\d{7,8}$/.test(dni) ? dni.padStart(8, '0') : null
        }
        const toNumber = (value: ExcelJS.CellValue): number => {
          if (typeof value === 'number') return value
          if (typeof value === 'object' && value !== null && 'result' in value) return toNumber(value.result as ExcelJS.CellValue)
          const number = Number(String(value ?? '').replace(',', '.'))
          return Number.isFinite(number) ? number : 0
        }
        const worksheet = workbook.getWorksheet('dias mes')
        if (!worksheet) throw new Error('No se encontró la hoja "dias mes".')

        const headers: string[] = []
        const columnMap = new Map<string, number>()
        worksheet.getRow(5).eachCell((cell, column) => {
          const header = normalizeHeader(cell.value)
          headers.push(header)
          if (header) columnMap.set(header, column)
        })

        console.log('ENCABEZADOS FILA 5:', headers)
        console.log('MAPA DE COLUMNAS:', Object.fromEntries(columnMap))

        const requiredColumns = ['DNI', ...months.map((month) => month.toUpperCase()), 'TOTAL']
        const missingColumns = requiredColumns.filter((column) => !columnMap.has(column))
        if (missingColumns.length > 0) {
          console.error('COLUMNAS FALTANTES:', missingColumns)
          throw new Error(`No se encontraron las columnas requeridas: ${missingColumns.join(', ')}.`)
        }

        const dniColumn = columnMap.get('DNI')!

        const daysByDni = new Map<string, MonthlyValues>()
        let processed = 0
        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber < 6) return
          const dni = normalizeDni(row.getCell(dniColumn).value)
          if (!dni) return
          processed += 1
          const monthValues = Object.fromEntries(months.map((month) => [month, toNumber(row.getCell(columnMap.get(month.toUpperCase())!).value)])) as Omit<MonthlyValues, 'total'>
          const calculatedTotal = months.reduce((sum, month) => sum + monthValues[month], 0)
          daysByDni.set(dni, { ...monthValues, total: columnMap.has('TOTAL') ? toNumber(row.getCell(columnMap.get('TOTAL')!).value) : calculatedTotal })
        })

        const noEncontrados = [...daysByDni.keys()].filter((dni) => !trabajadores.some((trabajador) => trabajador.dni === dni))
        const coincidentes = daysByDni.size - noEncontrados.length
        setDiasImportados(coincidentes)
        const primerTrabajador = trabajadores.find((trabajador) => daysByDni.has(trabajador.dni))

        setTrabajadores((current) =>
  current.map((trabajador) => {
    const diasTrabajados = daysByDni.get(trabajador.dni)

    return diasTrabajados
      ? {
          ...trabajador,
          diasTrabajados,
        }
      : trabajador
  })
)

setTrabajadores((current) =>
  current.map((trabajador) => {
    const diasTrabajados = daysByDni.get(trabajador.dni)

    return diasTrabajados
      ? {
          ...trabajador,
          diasTrabajados,
        }
      : trabajador
  })
)

        
        setImportedFile(file.name)

        console.log('DÍAS TRABAJADOS IMPORTADOS', [...daysByDni.entries()])
        console.log('TRABAJADORES COINCIDENTES', coincidentes)
        console.log('DNIs NO ENCONTRADOS', noEncontrados)
        console.log('PRIMER TRABAJADOR CON DÍAS', primerTrabajador ? { ...primerTrabajador, diasTrabajados: daysByDni.get(primerTrabajador.dni) } : undefined)

        alert(`Días trabajados importados correctamente.\n\nTrabajadores procesados: ${processed}\nDNI coincidentes: ${coincidentes}\nDNI no encontrados: ${noEncontrados.length}`)
        return
      }

      const worksheet = workbook.getWorksheet('REM')

      if (!worksheet) {
        alert('No se encontró la hoja "REM" en el archivo.')
        return
      }

      console.log('--- ENCABEZADOS HOJA REM ---')

worksheet.getRow(1).eachCell((cell, colNumber) => {
  console.log(colNumber, cell.value)
})

worksheet.getRow(2).eachCell((cell, colNumber) => {
  console.log('FILA 2:', colNumber, cell.value)
})

worksheet.getRow(3).eachCell((cell, colNumber) => {
  console.log('FILA 3:', colNumber, cell.value)
})

      const trabajadoresImportados: ImportedWorker[] = []

const getExcelNumber = (value: ExcelJS.CellValue): number => {
  if (typeof value === 'number') {
    return value
  }

  if (
    typeof value === 'object' &&
    value !== null &&
    'result' in value &&
    typeof value.result === 'number'
  ) {
    return value.result
  }

  const numero = Number(value)

  return Number.isFinite(numero) ? numero : 0
}


     worksheet.eachRow((row) => {
  const dni = row.getCell(2).value

  // Solo procesamos filas que realmente tengan un DNI válido
 const dniTexto = String(dni ?? '')
  .trim()
  .replace(/\.0$/, '')

if (!/^\d{7,8}$/.test(dniTexto)) {
  return
}

const dniFinal = dniTexto.padStart(8, '0')

  const apellidoPaterno = String(row.getCell(3).value ?? '').trim()
  const apellidoMaterno = String(row.getCell(4).value ?? '').trim()
  const nombres = String(row.getCell(5).value ?? '').trim()

  const fechaInicio = row.getCell(6).value
  const fechaCese = row.getCell(7).value

  trabajadoresImportados.push({
  dni: dniFinal,
  apellidoPaterno,
  apellidoMaterno,
  nombres,
  fechaInicio,
  fechaCese,

  remuneraciones: {
    enero: getExcelNumber(row.getCell(9).value),
    febrero: getExcelNumber(row.getCell(10).value),
    marzo: getExcelNumber(row.getCell(11).value),
    abril: getExcelNumber(row.getCell(12).value),
    mayo: getExcelNumber(row.getCell(13).value),
    junio: getExcelNumber(row.getCell(14).value),
    julio: getExcelNumber(row.getCell(15).value),
    agosto: getExcelNumber(row.getCell(16).value),
    septiembre: getExcelNumber(row.getCell(17).value),
    octubre: getExcelNumber(row.getCell(18).value),
    noviembre: getExcelNumber(row.getCell(19).value),
    diciembre: getExcelNumber(row.getCell(20).value),
    total: getExcelNumber(row.getCell(21).value),
  },
})
})

      console.log(
  'TRABAJADORES IMPORTADOS:',
  trabajadoresImportados
)

console.log(
  'REMUNERACIONES PRIMER TRABAJADOR:',
  trabajadoresImportados[0]?.remuneraciones
)

setTrabajadores(
  trabajadoresImportados.map((trabajador) => {
    const diasTrabajados = diasPendientes.get(trabajador.dni)

    return diasTrabajados
      ? {
          ...trabajador,
          diasTrabajados,
        }
      : trabajador
  })
)
setImportedFile(file.name)

alert(
  `Excel leído correctamente.\n\nTrabajadores encontrados: ${trabajadoresImportados.length}`
)
      } catch (error) {
       console.error('Error al leer Excel:', error)
       alert(`No se pudo leer el archivo Excel.${error instanceof Error ? `\n\n${error.message}` : ''}`)
    } finally {
      setLoading(false)
    }
  }

  const handleInputChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]

    if (file) {
      handleFile(file)
    }

    event.target.value = ''
  }

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
  {!importedFile ? (
    <>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={handleInputChange}
      />

      <button
        type="button"
        disabled={loading}
        onClick={() => inputRef.current?.click()}
        className="flex w-full flex-col items-center gap-3 rounded-2xl border border-dashed border-primary/40 bg-primary/5 px-6 py-9 text-center transition-colors hover:bg-primary/10 disabled:cursor-not-allowed disabled:opacity-60"
      >
        <FileSpreadsheet className="size-8 text-primary" />

        <span className="font-medium">
          {loading ? 'Leyendo archivo...' : 'Arrastra el archivo aquí'}
        </span>

        <span className="text-sm text-muted-foreground">
          o selecciónalo desde tu dispositivo
        </span>

        <span className="rounded-lg border bg-background px-3 py-2 text-sm font-medium">
          {loading ? 'Procesando...' : 'Seleccionar archivo'}
        </span>
      </button>
    </>
  ) : (
    <div className="space-y-4">

      {/* Archivo cargado */}
      <div className="flex items-center justify-between rounded-2xl bg-muted/30 p-4">
        <div className="flex items-center gap-3">
          <div className="flex size-10 items-center justify-center rounded-xl bg-primary/15">
            <FileSpreadsheet className="size-5 text-primary" />
          </div>

          <div>
            <p className="font-medium">
              {importedFile}
            </p>

            <p className="text-sm text-muted-foreground">
              {isRemunerations ? trabajadores.length : diasPendientes.size} trabajadores detectados · Mapeo automático disponible
            </p>
          </div>
        </div>

        <button
          type="button"
          onClick={() => inputRef.current?.click()}
          className="rounded-lg border px-3 py-2 text-sm font-medium transition-colors hover:bg-muted"
        >
          Cambiar archivo
        </button>
      </div>

      {/* Resumen */}
      <div className="grid gap-3 md:grid-cols-3">

        <div className="rounded-2xl border p-4">
          <p className="text-sm text-muted-foreground">
            Trabajadores detectados
          </p>

          <p className="mt-2 text-2xl font-semibold">
  {isRemunerations ? trabajadores.length : diasImportados}
</p>
        </div>

        <div className="rounded-2xl border p-4">
          <p className="text-sm text-muted-foreground">
            Remuneración anual
          </p>

          <p className="text-2xl font-semibold">
  {isRemunerations && trabajadores.length > 0
    ? `S/ ${trabajadores
        .reduce((total, trabajador) => {
          return total + trabajador.remuneraciones.total
        }, 0)
        .toLocaleString('es-PE', {
          minimumFractionDigits: 2,
          maximumFractionDigits: 2,
        })}`
    : '—'}
</p>
        </div>

        <div className="rounded-2xl border p-4">
          <p className="text-sm text-muted-foreground">
            Observaciones
          </p>

          <p className="mt-2 text-2xl font-semibold">
            0
          </p>
        </div>

      </div>

{/* Trabajadores */}
<div className="overflow-hidden rounded-2xl border">

  <div className="grid grid-cols-4 gap-4 border-b bg-muted/20 px-4 py-3 text-sm font-medium">
    <span>Trabajador</span>
    <span>DNI</span>
    <span>Inicio</span>
    <span>Cese</span>
  </div>

  <div className="divide-y">

    {trabajadores.map((trabajador, index) => (
      <div key={trabajador.dni}>

        {/* FILA DEL TRABAJADOR */}
        <div
          onClick={() =>
            setTrabajadorSeleccionado(
              trabajadorSeleccionado === index ? null : index
            )
          }
          className="grid grid-cols-4 gap-4 px-4 py-3 text-sm cursor-pointer hover:bg-muted/30"
        >

          <div>
            <p className="font-medium">
              {trabajador.apellidoPaterno}{" "}
              {trabajador.apellidoMaterno}
            </p>

            <p className="text-xs text-muted-foreground">
              {trabajador.nombres}
            </p>
          </div>

          <span>
            {trabajador.dni}
          </span>

          <span>
            {formatExcelDate(trabajador.fechaInicio)}
          </span>

          <span>
            {formatExcelDate(trabajador.fechaCese)}
          </span>

        </div>

        {/* DETALLE DEL TRABAJADOR */}
        {trabajadorSeleccionado === index && (
          <div className="border-t bg-muted/10 px-4 py-4">

            <p className="mb-3 text-sm font-semibold">
              Remuneraciones mensuales
            </p>

            <div className="grid grid-cols-2 gap-3 md:grid-cols-4">

              {[
                ["Enero", trabajador.remuneraciones.enero],
                ["Febrero", trabajador.remuneraciones.febrero],
                ["Marzo", trabajador.remuneraciones.marzo],
                ["Abril", trabajador.remuneraciones.abril],
                ["Mayo", trabajador.remuneraciones.mayo],
                ["Junio", trabajador.remuneraciones.junio],
                ["Julio", trabajador.remuneraciones.julio],
                ["Agosto", trabajador.remuneraciones.agosto],
                ["Septiembre", trabajador.remuneraciones.septiembre],
                ["Octubre", trabajador.remuneraciones.octubre],
                ["Noviembre", trabajador.remuneraciones.noviembre],
                ["Diciembre", trabajador.remuneraciones.diciembre],
              ].map(([mes, monto]) => (
                <div
                  key={mes}
                  className="rounded-lg border border-border/60 bg-background p-3"
                >

                  <p className="text-xs text-muted-foreground">
                    {mes}
                  </p>

                  <p className="mt-1 text-sm font-semibold">
                    S/{" "}
                    {Number(monto).toLocaleString("es-PE", {
                      minimumFractionDigits: 2,
                      maximumFractionDigits: 2,
                    })}
                  </p>

                </div>
              ))}

            </div>

            <div className="mt-3 flex items-center justify-between rounded-lg border border-primary/20 bg-primary/5 p-3">

              <span className="text-sm font-medium">
                Remuneración total anual
              </span>

              <span className="font-semibold text-primary">
                S/{" "}
                {Number(trabajador.remuneraciones.total).toLocaleString(
                  "es-PE",
                  {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2,
                  }
                )}
              </span>

            </div>

            {trabajador.diasTrabajados && (
              <div className="mt-4">
                <p className="mb-3 text-sm font-semibold">Días trabajados</p>
                <div className="grid grid-cols-2 gap-3 md:grid-cols-4">
                  {[
                    ['Enero', trabajador.diasTrabajados.enero], ['Febrero', trabajador.diasTrabajados.febrero],
                    ['Marzo', trabajador.diasTrabajados.marzo], ['Abril', trabajador.diasTrabajados.abril],
                    ['Mayo', trabajador.diasTrabajados.mayo], ['Junio', trabajador.diasTrabajados.junio],
                    ['Julio', trabajador.diasTrabajados.julio], ['Agosto', trabajador.diasTrabajados.agosto],
                    ['Septiembre', trabajador.diasTrabajados.septiembre], ['Octubre', trabajador.diasTrabajados.octubre],
                    ['Noviembre', trabajador.diasTrabajados.noviembre], ['Diciembre', trabajador.diasTrabajados.diciembre],
                  ].map(([mes, dias]) => (
                    <div key={mes} className="rounded-lg border border-border/60 bg-background p-3">
                      <p className="text-xs text-muted-foreground">{mes}</p>
                      <p className="mt-1 text-sm font-semibold">{dias} días</p>
                    </div>
                  ))}
                </div>
                <div className="mt-3 flex items-center justify-between rounded-lg border border-primary/20 bg-primary/5 p-3">
                  <span className="text-sm font-medium">Total de días trabajados</span>
                  <span className="font-semibold text-primary">{trabajador.diasTrabajados.total} días</span>
                </div>
              </div>
            )}

          </div>
        )}

      </div>
    ))}

  </div>

</div>

    </div>
  )}
</CardContent>
    </Card>
  )
}
export function CalculationView({ results, totals, errors, onRun, onSelect }: { results: EmployeeUtilityResult[]; totals: ReturnType<typeof getTotals>; errors: string[]; onRun: () => void; onSelect: (result: EmployeeUtilityResult) => void }) { const top = [...results].sort((a,b) => b.finalAmount - a.finalAmount).slice(0,5);
 const [trabajadores, setTrabajadores] = useState<ImportedWorker[]>([])
; return (
<div className="flex flex-col gap-7">
  <SectionTitle eyebrow="Cálculo · Paso 2 de 3" title="Distribución de utilidades" description="Revisa los parámetros, valida los factores de distribución y genera el cálculo anual del ejercicio." action={<div className="flex gap-2 sm:hidden">
  <Button variant="outline" size="sm">
  <Check className="mr-2 size-4" />Guardar</Button>
<Button size="sm" onClick={onRun}>
  <Play className="mr-2 size-4" />Ejecutar</Button></div>} />
<Card className="border-primary/30 bg-primary/5">
  <CardContent className="flex flex-col gap-4 p-5 sm:flex-row sm:items-center">
  <div className="flex size-11 items-center justify-center rounded-xl bg-primary/15 text-primary"><Upload /></div>
<div className="flex-1">
  <p className="font-semibold">Base de trabajadores lista para procesar</p>
<p className="mt-1 text-sm text-muted-foreground">27 registros · Última importación 18 ago 2026 · Excel validado</p></div>
<Button variant="outline" onClick={onRun}>Procesar ahora <ArrowRight data-icon="inline-end" /></Button></CardContent></Card>
{/* Importación de información para el cálculo */}
<div className="grid gap-6 xl:grid-cols-2">
<ImportCard
   kind="remunerations"
   trabajadores={trabajadores}
   setTrabajadores={setTrabajadores}
 />

<ImportCard
   kind="worked-days"
   trabajadores={trabajadores}
   setTrabajadores={setTrabajadores}
 />
</div>

<div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
<Kpi
  label="Fondo legal"
  value={money.format(totals.fund)}
  hint="10% de la renta neta"
  icon={ShieldCheck}
  tone="teal"
/>

<Kpi
  label="A distribuir"
  value={money.format(totals.distributed)}
  hint="50% días · 50% remuneración"
  icon={ArrowDownToLine}
/>

<Kpi
  label="Trabajadores"
  value={number.format(results.length)}
  hint={`${results.length} registros procesados`}
  icon={Users}
  tone="amber"
/>

<Kpi
  label="Remanente"
  value={money.format(totals.remainder)}
  hint={`${totals.capped} topes aplicados`}
  icon={AlertCircle}
  tone="rose"
/>
</div>

{errors.length > 0 && (
  <Alert variant="destructive">
    <AlertCircle className="size-4" />

    <AlertTitle>
      Validaciones pendientes
    </AlertTitle>

    <AlertDescription>
      {errors.join(' ')}
    </AlertDescription>
  </Alert>
)}

<div className="grid gap-5 xl:grid-cols-[1.15fr_.85fr]">

  {/* RESUMEN DEL FONDO */}
  <Card>
    <CardHeader>
      <div className="flex items-start justify-between gap-4">
        <div>
          <CardTitle>
            Resumen del fondo
          </CardTitle>

          <CardDescription>
            Distribución legal aplicable al ejercicio 2026
          </CardDescription>
        </div>

        <Badge
          variant="secondary"
          className="bg-emerald-500/10 text-emerald-400"
        >
          Validado
        </Badge>
      </div>
    </CardHeader>

    <CardContent>
      <div className="flex flex-col gap-5">

        <div className="grid gap-4 sm:grid-cols-3">

          <div>
            <p className="text-xs text-muted-foreground">
              Renta neta
            </p>

            <p className="mt-1 text-lg font-semibold">
              {money.format(parameters.income)}
            </p>
          </div>

          <div>
            <p className="text-xs text-muted-foreground">
              Porcentaje legal
            </p>

            <p className="mt-1 text-lg font-semibold">
              {parameters.legalPercent}%
            </p>
          </div>

          <div>
            <p className="text-xs text-muted-foreground">
              Fondo total
            </p>

            <p className="mt-1 text-lg font-semibold text-primary">
              {money.format(calculateFund(parameters))}
            </p>
          </div>

        </div>

        <Separator />

        <div>
          <div className="mb-2 flex justify-between text-sm">
            <span>
              Distribuido
            </span>

            <span className="font-medium">
              {((totals.distributed / totals.fund) * 100).toFixed(1)}%
            </span>
          </div>

          <Progress
            value={(totals.distributed / totals.fund) * 100}
            className="h-2"
          />

          <div className="mt-2 flex justify-between text-xs text-muted-foreground">
            <span>
              {money.format(totals.distributed)}
            </span>

            <span>
              Remanente {money.format(totals.remainder)}
            </span>
          </div>
        </div>

        <div className="rounded-lg border border-border/70 bg-muted/30 p-4">
          <div className="flex items-start gap-3">

            <div className="rounded-md bg-primary/10 p-2 text-primary">
              <Calculator className="size-4" />
            </div>

            <div>
              <p className="text-sm font-medium">
                Fórmula de distribución
              </p>

              <p className="mt-1 text-xs leading-5 text-muted-foreground">
                50% se asigna según días laborados y 50% según
                remuneración computable. Cada resultado está sujeto
                al tope de 18 remuneraciones.
              </p>
            </div>

          </div>
        </div>

      </div>
    </CardContent>
  </Card>

  {/* TOP 5 RESULTADOS */}
  <Card>
    <CardHeader>
      <CardTitle>
        Top 5 resultados
      </CardTitle>

      <CardDescription>
        Mayor monto individual calculado
      </CardDescription>
    </CardHeader>

    <CardContent>
      <div className="flex flex-col gap-4">

        {top.map((result, index) => (
          <button
            key={result.id}
            onClick={() => onSelect(result)}
            className="flex items-center gap-3 text-left"
          >
            <span className="flex size-6 items-center justify-center rounded-full bg-muted text-xs font-semibold">
              {index + 1}
            </span>

            <div className="min-w-0 flex-1">

              <p className="truncate text-sm font-medium">
                {result.name}
              </p>

              <div className="mt-1 h-1.5 overflow-hidden rounded-full bg-muted">
                <div
                  className="h-full rounded-full bg-primary"
                  style={{
                    width: `${Math.max(
                      8,
                      (result.finalAmount / top[0].finalAmount) * 100
                    )}%`,
                  }}
                />
              </div>

            </div>

            <span className="text-sm font-semibold">
              {money.format(result.finalAmount)}
            </span>
          </button>
        ))}

      </div>
    </CardContent>
  </Card>

</div>

{/* DISTRIBUCIÓN POR TRABAJADOR */}
<Card>
  <CardHeader>

    <div className="flex items-center justify-between">

      <div>
        <CardTitle>
          Distribución por trabajador
        </CardTitle>

        <CardDescription>
          Vista previa de los resultados calculados
        </CardDescription>
      </div>

      <Button
        variant="outline"
        size="sm"
      >
        <Download className="mr-2 size-4" />
        Exportar
      </Button>

    </div>

  </CardHeader>

  <CardContent>

    <div className="overflow-x-auto">

      <Table>

        <TableHeader>
          <TableRow>

            <TableHead>
              Trabajador
            </TableHead>

            <TableHead>
              Días
            </TableHead>

            <TableHead>
              Remuneración
            </TableHead>

            <TableHead className="text-right">
              Preliminar
            </TableHead>

            <TableHead className="text-right">
              Final
            </TableHead>

          </TableRow>
        </TableHeader>

        <TableBody>

          {results.slice(0, 6).map((result) => (
            <TableRow
              key={result.id}
              className="cursor-pointer"
              onClick={() => onSelect(result)}
            >

              <TableCell>

                <div>

                  <p className="font-medium">
                    {result.name}
                  </p>

                  <p className="text-xs text-muted-foreground">
                    {result.code} · {result.department}
                  </p>

                </div>

              </TableCell>

              <TableCell>
                {number.format(result.days)}
              </TableCell>

              <TableCell>
                {money.format(result.remuneration)}
              </TableCell>

              <TableCell className="text-right">
                {money.format(result.preliminary)}
              </TableCell>

              <TableCell className="text-right font-semibold text-primary">
                {money.format(result.finalAmount)}
              </TableCell>

            </TableRow>
          ))}

        </TableBody>

      </Table>

    </div>

  </CardContent>
</Card>
</div>
)
}
