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

  type IncidenciaImportada = {
  dni: string
  codigo: string
  descripcion: string
  cantidadDias: number
  fechaInicio: ExcelJS.CellValue
  fechaFin: ExcelJS.CellValue
}

type FeriadoImportado = {
  fecha: ExcelJS.CellValue
  descripcion: string
}

type ImportedWorker = {
  dni: string
  apellidoPaterno: string
  apellidoMaterno: string
  nombres: string

  // Datos laborales originales encontrados en los archivos
  fechaInicio: ExcelJS.CellValue
  fechaCese: ExcelJS.CellValue

  // Información consolidada de remuneraciones
  remuneraciones: MonthlyValues

  // Información consolidada de días
  diasTrabajados?: MonthlyValues

  // Meses en los que el trabajador fue encontrado
  mesesRemuneraciones?: string[]
  mesesDias?: string[]

  // Meses en los que existe información en cualquiera de las fuentes
  mesesPresentes?: string[]

  // Información para las siguientes etapas
  incidencias?: IncidenciaImportada[]
  diasNoLaborados?: number
  diasEfectivos?: number
}

  const months = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']   as const

  const MESES_EN_CERO: MonthlyValues = { enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0, julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 0 }

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

    const [importedFile, setImportedFile] = useState<string | null>(null)
  const [trabajadorSeleccionado, setTrabajadorSeleccionado] = useState<number | null>(null)
  const formatDate = (date: Date) => {
    const day = String(date.getUTCDate()).padStart(2, '0')
    const month = String(date.getUTCMonth() + 1).padStart(2, '0')
    const year = date.getUTCFullYear()

    return `${day}/${month}/${year}`
  }

  const formatExcelDate = (value: ExcelJS.CellValue) => {
    if (!value) return '—'

    if (value instanceof Date) {
      return formatDate(value)
    }

    if (typeof value === 'object') {
      if ('result' in value) {
        return formatExcelDate(value.result as ExcelJS.CellValue)
      }

      if ('text' in value && typeof value.text === 'string') {
        return formatExcelDate(value.text)
      }
    }

    if (typeof value === 'number') {
      const excelEpoch = Date.UTC(1899, 11, 30)
      return formatDate(new Date(excelEpoch + value * 86_400_000))
    }

    if (typeof value === 'string') {
      const date = new Date(value)
      return Number.isNaN(date.getTime()) ? value : formatDate(date)
    }

    return '—'
  }

    const handleFile = async (file: File) => {
      setLoading(true)

      try {
        const workbook = new ExcelJS.Workbook()
        const buffer = await file.arrayBuffer()

        await workbook.xlsx.load(buffer)

        if (!isRemunerations) {
  const normalizeHeader = (value: ExcelJS.CellValue) => {
    const text =
      typeof value === 'object' &&
      value !== null &&
      'text' in value
        ? String(value.text ?? '')
        : String(value ?? '')

    const normalized = text
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .trim()
      .replace(/\s+/g, ' ')
      .toUpperCase()

    return normalized === 'SETIEMBRE' ? 'SEPTIEMBRE' : normalized
  }

  const normalizeDni = (value: ExcelJS.CellValue) => {
    const dni = String(value ?? '')
      .trim()
      .replace(/\.0$/, '')

    return /^\d{7,8}$/.test(dni)
      ? dni.padStart(8, '0')
      : null
  }

  const toNumber = (value: ExcelJS.CellValue): number => {
    if (typeof value === 'number') {
      return value
    }

    if (
      typeof value === 'object' &&
      value !== null &&
      'result' in value
    ) {
      return toNumber(value.result as ExcelJS.CellValue)
    }

    const numero = Number(
      String(value ?? '').replace(',', '.')
    )

    return Number.isFinite(numero) ? numero : 0
  }

  // ============================================================
  // PASO 1: VERIFICAR LAS HOJAS DEL EXCEL
  // ============================================================

  const nombresHojas = workbook.worksheets.map(
    (sheet) => sheet.name
  )

  console.log('======================================')
  console.log('EXCEL DE DÍAS CARGADO')
  console.log('HOJAS ENCONTRADAS:')
  console.log(nombresHojas)
  console.log('======================================')

  // ============================================================
  // PASO 2: BUSCAR HOJA "dias mes"
  // ============================================================

  const worksheetDiasMes =
    workbook.getWorksheet('dias mes')

  if (!worksheetDiasMes) {
    throw new Error(
      `No se encontró la hoja "dias mes".

Hojas encontradas:
${nombresHojas.join(', ')}`
    )
  }

  console.log('HOJA "dias mes" ENCONTRADA')

  // ============================================================
  // PASO 3: BUSCAR HOJA "DIAS"
  // ============================================================

  const worksheetDias =
    workbook.getWorksheet('DIAS')

  if (worksheetDias) {
    console.log('HOJA "DIAS" ENCONTRADA')
  } else {
    console.warn(
      'ADVERTENCIA: No se encontró la hoja "DIAS".'
    )
  }

  // ============================================================
  // PASO 4: MOSTRAR LAS COLUMNAS DE "dias mes"
  // ============================================================

  console.log('--------------------------------------')
  console.log('COLUMNAS DE "dias mes"')
  console.log('--------------------------------------')

  const encabezadosDiasMes: {
    columna: number
    original: string
    normalizado: string
  }[] = []

  worksheetDiasMes.getRow(5).eachCell(
    (cell, column) => {
      const original = String(
        cell.value ?? ''
      ).trim()

      const normalizado =
        normalizeHeader(cell.value)

      if (original) {
        encabezadosDiasMes.push({
          columna: column,
          original,
          normalizado,
        })

        console.log(
          `Columna ${column}: "${original}" → "${normalizado}"`
        )
      }
    }
  )

  // ============================================================
  // PASO 5: VERIFICAR COLUMNAS NECESARIAS
  // ============================================================

  const columnMap = new Map<string, number>()

  encabezadosDiasMes.forEach(
    ({ columna, normalizado }) => {
      if (normalizado) {
        columnMap.set(
          normalizado,
          columna
        )
      }
    }
  )

  const requiredColumns = [
    'DNI',
    ...months.map((month) =>
      month.toUpperCase()
    ),
    'TOTAL',
  ]

  const missingColumns =
    requiredColumns.filter(
      (column) =>
        !columnMap.has(column)
    )

  if (missingColumns.length > 0) {
    console.error(
      'COLUMNAS FALTANTES EN "dias mes":',
      missingColumns
    )

    throw new Error(
      `La hoja "dias mes" no tiene las columnas requeridas:

${missingColumns.join(', ')}`
    )
  }

  console.log(
    'COLUMNAS NECESARIAS ENCONTRADAS CORRECTAMENTE'
  )

  // ============================================================
  // PASO 6: EXTRAER LOS DÍAS DE "dias mes"
  // ============================================================

  const dniColumn =
    columnMap.get('DNI')!

  const daysByDni =
    new Map<string, MonthlyValues>()

  let processed = 0

  worksheetDiasMes.eachRow(
    (row, rowNumber) => {
      if (rowNumber < 6) {
        return
      }

      const dni =
        normalizeDni(
          row.getCell(
            dniColumn
          ).value
        )

      if (!dni) {
        return
      }

      processed += 1

      const monthValues =
        Object.fromEntries(
          months.map(
            (month) => [
              month,
              toNumber(
                row.getCell(
                  columnMap.get(
                    month.toUpperCase()
                  )!
                ).value
              ),
            ]
          )
        ) as Omit<
          MonthlyValues,
          'total'
        >

      const calculatedTotal =
        months.reduce(
          (sum, month) =>
            sum +
            monthValues[month],
          0
        )

      const total =
        columnMap.has('TOTAL')
          ? toNumber(
              row.getCell(
                columnMap.get(
                  'TOTAL'
                )!
              ).value
            )
          : calculatedTotal

      daysByDni.set(
        dni,
        {
          ...monthValues,
          total,
        }
      )
    }
  )

  console.log(
    'TRABAJADORES CON DÍAS EXTRAÍDOS:',
    daysByDni.size
  )

  const nombresPorDni =
    new Map<
      string,
      {
        apellidoPaterno: string
        apellidoMaterno: string
        nombres: string
      }
    >()


  // ============================================================
// PASO 7: EXTRAER INCIDENCIAS Y FERIADOS DE "DIAS"
// ============================================================

const incidenciasPorDni =
  new Map<string, IncidenciaImportada[]>()

const feriadosImportados: FeriadoImportado[] = []

if (worksheetDias) {
  console.log('======================================')
  console.log('ANALIZANDO HOJA "DIAS"')
  console.log('======================================')


  // ============================================================
// PASO 7: EXTRAER INCIDENCIAS DE "DIAS"
// ============================================================

// La hoja "DIAS" es un reporte PLAME.
// Los encabezados están distribuidos en las filas 9 y 10.
// Los datos comienzan en la fila 11.
//
// Estructura detectada:
// A = Mes
// B = DNI
// C = Apellido paterno
// D = Apellido materno
// E = Nombres
// F = Código
// G-I = Descripción
// J = Num. Días
// K = Otro dato del reporte
// L = Fecha inicial
// M = Fecha final
// N = Otro dato del reporte

const incidenciasPorDni =
  new Map<string, IncidenciaImportada[]>()

const feriadosImportados: FeriadoImportado[] = []
if (worksheetDias) {
  console.log('======================================')
  console.log('EXTRAYENDO INCIDENCIAS DE "DIAS"')
  console.log('======================================')
  const FILA_DATOS = 11
  const COL_DNI = 2
  const COL_CODIGO = 6
  const COL_DESCRIPCION = 7
  const COL_CANTIDAD_DIAS = 10
  const COL_FECHA_INICIO = 12
  const COL_FECHA_FIN = 13
  let totalIncidencias = 0
  worksheetDias.eachRow(
    (row, rowNumber) => {
      if (rowNumber < FILA_DATOS) {
        return
      }
      // ----------------------------------------------
      // DNI
      // ----------------------------------------------
      const dni =
        normalizeDni(
          row.getCell(
            COL_DNI
          ).value
        )
      if (!dni) {
        return
      }
      // ----------------------------------------------
      // Nombres (columnas C / D / E)
      // ----------------------------------------------
      const apellidoPaterno =
        String(row.getCell(3).value ?? '').trim()
      const apellidoMaterno =
        String(row.getCell(4).value ?? '').trim()
      const nombres =
        String(row.getCell(5).value ?? '').trim()
      const nombresPrevios =
        nombresPorDni.get(dni)
      nombresPorDni.set(
        dni,
        {
          apellidoPaterno:
            apellidoPaterno ||
            nombresPrevios?.apellidoPaterno ||
            '',
          apellidoMaterno:
            apellidoMaterno ||
            nombresPrevios?.apellidoMaterno ||
            '',
          nombres:
            nombres ||
            nombresPrevios?.nombres ||
            '',
        }
      )
      // ----------------------------------------------
      // CÓDIGO
      // ----------------------------------------------
      const codigo =
        String(
          row.getCell(
            COL_CODIGO
          ).value ?? ''
        ).trim()
      // ----------------------------------------------
      // DESCRIPCIÓN
      // ----------------------------------------------
      const descripcion =
        String(
          row.getCell(
            COL_DESCRIPCION
          ).value ?? ''
        ).trim()
      // ----------------------------------------------
      // CANTIDAD DE DÍAS
      // ----------------------------------------------
      const cantidadDias =
        toNumber(
          row.getCell(
            COL_CANTIDAD_DIAS
          ).value
        )
      // ----------------------------------------------
      // FECHA INICIAL
      // ----------------------------------------------
      const fechaInicio =
        row.getCell(
          COL_FECHA_INICIO
        ).value
      // ----------------------------------------------
      // FECHA FINAL
      // ----------------------------------------------
      const fechaFin =
        row.getCell(
          COL_FECHA_FIN
        ).value
      // ----------------------------------------------
      // Ignorar filas sin incidencia
      // ----------------------------------------------
      if (
        !descripcion &&
        cantidadDias === 0 &&
        !fechaInicio &&
        !fechaFin
      ) {
        return
      }
      const incidencia: IncidenciaImportada =
        {
          dni,
          codigo,
          descripcion,
          cantidadDias,
          fechaInicio,
          fechaFin,
        }
      const incidenciasExistentes =
        incidenciasPorDni.get(
          dni
        ) ?? []
      incidenciasExistentes.push(
        incidencia
      )
      incidenciasPorDni.set(
        dni,
        incidenciasExistentes
      )
      totalIncidencias += 1
    }
  )
  console.log(
    'INCIDENCIAS EXTRAÍDAS:',
    totalIncidencias
  )
  console.log(
    'TRABAJADORES CON INCIDENCIAS:',
    incidenciasPorDni.size
  )
  // ----------------------------------------------
  // Mostrar las primeras incidencias
  // para comprobar que están bien extraídas
  // ----------------------------------------------
  let contador = 0
  for (
    const [
      dni,
      incidencias,
    ] of incidenciasPorDni
  ) {
    console.log(
      `DNI ${dni}:`,
      incidencias
    )
    contador += 1
    if (contador >= 10) {
      break
    }
  }
}
  // ==========================================================
  // BUSCAR FERIADOS
  // ==========================================================
  console.log(
    '======================================'
  )
  console.log(
    'BUSCANDO FERIADOS EN "DIAS"...'
  )
  let feriadosHeaderRow: number | null =
    null
  let feriadosHeaderMap =
    new Map<string, number>()
  worksheetDias.eachRow(
    (row, rowNumber) => {
      if (
        feriadosHeaderRow !== null
      ) {
        return
      }
      const posiblesHeaders =
        new Map<string, number>()
      row.eachCell(
        (cell, column) => {
          const header =
            normalizeHeader(
              cell.value
            )
          if (header) {
            posiblesHeaders.set(
              header,
              column
            )
          }
        }
      )
      const tieneFecha =
        posiblesHeaders.has(
          'FECHA'
        )
      const tieneDescripcion =
        posiblesHeaders.has(
          'DESCRIPCION'
        ) ||
        posiblesHeaders.has(
          'FERIADO'
        )
      if (
        tieneFecha &&
        tieneDescripcion
      ) {
        feriadosHeaderRow =
          rowNumber

        feriadosHeaderMap =
          posiblesHeaders
      }
    }
  )
  if (
    feriadosHeaderRow !== null
  ) {
    const fechaColumn =
      feriadosHeaderMap.get(
        'FECHA'
      )
    const descripcionColumn =
      feriadosHeaderMap.get(
        'DESCRIPCION'
      ) ??
      feriadosHeaderMap.get(
        'FERIADO'
      )
    worksheetDias.eachRow(
      (row, rowNumber) => {
        if (
          rowNumber <=
          feriadosHeaderRow!
        ) {
          return
        }
        if (!fechaColumn) {
          return
        }
        const fecha =
          row.getCell(
            fechaColumn
          ).value
        if (!fecha) {
          return
        }
        const descripcion =
          descripcionColumn
            ? String(
                row.getCell(
                  descripcionColumn
                ).value ?? ''
              ).trim()
            : ''
        feriadosImportados.push({
          fecha,
          descripcion,
        })
      }
    )

    console.log(
      'FERIADOS EXTRAÍDOS:',
      feriadosImportados.length
    )

    console.log(
      'FERIADOS:',
      feriadosImportados
    )
  } else {
    console.warn(
      'No se encontró automáticamente una tabla de feriados en "DIAS".'
    )
  }

  // ==========================================================
  // GUARDAR INCIDENCIAS EN CADA TRABAJADOR
  // ==========================================================
  setTrabajadores(
    (current) =>
      current.map(
        (trabajador) => {
          const incidencias =
            incidenciasPorDni.get(
              trabajador.dni
            )
          return {
            ...trabajador,
            incidencias:
              incidencias ?? [],
          }
        }
      )
  )
  console.log(
    '======================================'
  )
  console.log(
    'EXTRACCIÓN DE DIAS TERMINADA'
  )
  console.log(
    'Días mensuales:',
    daysByDni.size
  )
  console.log(
    'Trabajadores con incidencias:',
    incidenciasPorDni.size
  )
  console.log(
    'Feriados:',
    feriadosImportados.length
  )
  console.log(
    '======================================'
  )
}
  // ============================================================
  // PASO 8: COMPARAR LOS DNI CON LOS TRABAJADORES
  // ============================================================
  const noEncontrados =
    [
      ...daysByDni.keys(),
    ].filter(
      (dni) =>
        !trabajadores.some(
          (trabajador) =>
            trabajador.dni === dni
        )
    )

  const coincidentes =
    daysByDni.size -
    noEncontrados.length

  setDiasImportados(
  daysByDni.size
)

  // ============================================================
  // PASO 9: GUARDAR LOS DÍAS EXTRAÍDOS
  // ============================================================

  setTrabajadores(
    (current) => {
      const actualizados = current.map(
        (trabajador) => {
          const clave =
  String(trabajador.dni).trim().padStart(8, '0')
  console.log(
  '[REM→DÍAS] Estado recibido:',
  {
    dni: clave,
    tieneRemuneraciones: !!trabajador.remuneraciones,
    mesesRemuneraciones: trabajador.mesesRemuneraciones,
  }
)

          const diasTrabajados =
  daysByDni.get(clave)

          const infoNombres =
  nombresPorDni.get(clave)

          if (!diasTrabajados && !infoNombres) {
            return trabajador
          }


          console.log(
  '[TRAZABILIDAD] Trabajador antes de actualizar DÍAS:',
  trabajador.dni,
  {
    mesesRemuneraciones: trabajador.mesesRemuneraciones,
    mesesDias: trabajador.mesesDias,
  }
)
         const mesesDias = diasTrabajados
  ? months.filter((mes) => diasTrabajados[mes] > 0)
  : trabajador.mesesDias ?? []

const mesesPresentes = Array.from(
  new Set([
    ...(trabajador.mesesRemuneraciones ?? []),
    ...mesesDias,
  ])
)

console.log(
  '[TRAZABILIDAD] Trabajador después de consolidar:',
  trabajador.dni,
  {
    mesesRemuneraciones: trabajador.mesesRemuneraciones,
    mesesDias,
    mesesPresentes,
  }
)

return {
  ...trabajador,
  apellidoPaterno:
    trabajador.apellidoPaterno ||
    infoNombres?.apellidoPaterno ||
    '',
  apellidoMaterno:
    trabajador.apellidoMaterno ||
    infoNombres?.apellidoMaterno ||
    '',
  nombres:
    trabajador.nombres ||
    infoNombres?.nombres ||
    '',
  ...(diasTrabajados
    ? { diasTrabajados }
    : {}),
  mesesDias,
  mesesPresentes,
}
        }
      )

      const nuevos: ImportedWorker[] = [...daysByDni.keys()]
        .filter(
          (dni) =>
            !current.some(
              (trabajador) =>
                String(trabajador.dni).trim().padStart(8, '0') === dni
            )
        )
        .map(
          (dni) => {
            const infoNombres = nombresPorDni.get(dni)

            const diasTrabajados = daysByDni.get(dni)!

const mesesDias = months.filter(
  (mes) => diasTrabajados[mes] > 0
)

return {
  dni,
  apellidoPaterno: infoNombres?.apellidoPaterno ?? '',
  apellidoMaterno: infoNombres?.apellidoMaterno ?? '',
  nombres: infoNombres?.nombres ?? '',
  fechaInicio: null,
  fechaCese: null,
  remuneraciones: MESES_EN_CERO,
  diasTrabajados,
  mesesDias,
  mesesPresentes: mesesDias,
}
          }
        )

      return [...actualizados, ...nuevos]
    }
  )

  setImportedFile(
    file.name
  )

  // ============================================================
  // PASO 10: RESULTADO DE ESTA PRIMERA ETAPA
  // ============================================================

  return
}
const worksheet = workbook.getWorksheet('REM')
        if (!worksheet) {
          alert('No se encontró la hoja "REM" en el archivo.')
          return
        }

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


      worksheet.eachRow((row: ExcelJS.Row) => {
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

    const remuneraciones: MonthlyValues = {
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
}

// Detectar los meses que realmente tienen remuneración registrada
const mesesRemuneraciones = months.filter(
  (mes) => remuneraciones[mes] > 0
)

trabajadoresImportados.push({
  dni: dniFinal,
  apellidoPaterno,
  apellidoMaterno,
  nombres,
  fechaInicio,
  fechaCese,

  remuneraciones,

  mesesRemuneraciones,

  // En esta etapa solo existe información REM.
  // DÍAS se agregará cuando se importe el Excel de días.
  mesesPresentes: mesesRemuneraciones,
})
  })
  console.log(
  '[TRAZABILIDAD REM] Trabajador importado antes de consolidar:',
  trabajadoresImportados[0]?.dni,
  {
    remuneraciones: trabajadoresImportados[0]?.remuneraciones,
    mesesRemuneraciones: trabajadoresImportados[0]?.mesesRemuneraciones,
    mesesPresentes: trabajadoresImportados[0]?.mesesPresentes,
  }
)
  setTrabajadores((current) => {
  return trabajadoresImportados.map((rem) => {
    const existente = current.find(
      (trabajador) =>
        String(trabajador.dni).trim().padStart(8, '0') === rem.dni
    )

    if (!existente) {
      return rem
    }

    return {
      ...existente,
      ...rem,
      diasTrabajados: existente.diasTrabajados,
      mesesDias: existente.mesesDias,
      mesesPresentes: Array.from(
        new Set([
          ...(rem.mesesRemuneraciones ?? []),
          ...(existente.mesesDias ?? []),
        ])
      ),
    }
  })
})
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
                {isRemunerations ? trabajadores.length : diasImportados} trabajadores detectados · Mapeo automático disponible
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
    {isRemunerations ? 'Remuneración anual' : 'Días registrados'}
  </p>
  <p className="mt-2 text-2xl font-semibold">
    {isRemunerations && trabajadores.length > 0
      ? `S/ ${trabajadores
          .reduce((total, trabajador) => {
            return total + trabajador.remuneraciones.total
          }, 0)
          .toLocaleString('es-PE', {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          })}`
      : !isRemunerations && trabajadores.length > 0
        ? `${trabajadores.reduce((total, trabajador) => {
            return total + (trabajador.diasTrabajados?.total ?? 0)
          }, 0)} días`
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
    <div className={`grid gap-4 border-b bg-muted/20 px-4 py-3 text-sm font-medium ${isRemunerations ? 'grid-cols-4' : 'grid-cols-[minmax(0,1.5fr)_minmax(5.5rem,0.7fr)_minmax(6rem,0.8fr)_minmax(6rem,0.8fr)]'}`}>
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
            className={`grid gap-4 px-4 py-3 text-sm cursor-pointer hover:bg-muted/30 ${isRemunerations ? 'grid-cols-4' : 'grid-cols-[minmax(0,1.5fr)_minmax(5.5rem,0.7fr)_minmax(6rem,0.8fr)_minmax(6rem,0.8fr)]'}`}
          >
            <div className="min-w-0">
              <p className="font-medium">
                {trabajador.apellidoPaterno}{" "}
                {trabajador.apellidoMaterno}
              </p>
              <p className="text-xs text-muted-foreground">
                {trabajador.nombres}
              </p>
            </div>
            <span className="min-w-0 overflow-hidden text-ellipsis whitespace-nowrap">
              {trabajador.dni}
            </span>
            <span className="min-w-0 whitespace-nowrap">
              {formatExcelDate(trabajador.fechaInicio)}
            </span>
            <span className="min-w-0 whitespace-nowrap">
              {formatExcelDate(trabajador.fechaCese)}
            </span>
          </div>
          {/* DETALLE DEL TRABAJADOR */}
          {trabajadorSeleccionado === index && (
            <div className="border-t bg-muted/10 px-4 py-4">

              {isRemunerations && (
  <>
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
    </>
  )}
              {!isRemunerations && trabajador.diasTrabajados && (
  <div className="mt-5 rounded-xl border border-border/60 bg-muted/10 p-4">
    <div className="mb-4 flex items-center justify-between">
      <div>
        <p className="text-sm font-semibold">
          Días trabajados
        </p>
        <p className="text-xs text-muted-foreground">
          Resumen mensual registrado
        </p>
      </div>
      <div className="rounded-lg border border-primary/20 bg-primary/10 px-3 py-1.5">
        <span className="text-sm font-semibold text-primary">
          {trabajador.diasTrabajados.total} días
        </span>
      </div>
    </div>
    {/* ============================================================
    CONSOLIDADO DE MESES DEL TRABAJADOR
    ============================================================ */}
{(trabajador.mesesRemuneraciones ||
  trabajador.mesesDias ||
  trabajador.mesesPresentes) && (
  <div className="mt-5 rounded-xl border border-border/60 bg-muted/10 p-4">

    <div className="mb-4">
      <p className="text-sm font-semibold">
        Meses registrados
      </p>

      <p className="text-xs text-muted-foreground">
        Trazabilidad de la información encontrada en Remuneraciones y Días laborados
      </p>
    </div>

    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6">
      {[
        ['Enero', 'enero'],
        ['Febrero', 'febrero'],
        ['Marzo', 'marzo'],
        ['Abril', 'abril'],
        ['Mayo', 'mayo'],
        ['Junio', 'junio'],
        ['Julio', 'julio'],
        ['Agosto', 'agosto'],
        ['Septiembre', 'septiembre'],
        ['Octubre', 'octubre'],
        ['Noviembre', 'noviembre'],
        ['Diciembre', 'diciembre'],
      ].map(([nombre, mes]) => {

        const tieneRem =
          trabajador.mesesRemuneraciones?.includes(mes) ?? false

        const tieneDias =
          trabajador.mesesDias?.includes(mes) ?? false

        const tieneInformacion =
          tieneRem || tieneDias

        return (
          <div
            key={mes}
            className={`rounded-lg border p-3 ${
              tieneInformacion
                ? 'border-primary/20 bg-primary/5'
                : 'border-border/60 bg-background'
            }`}
          >
            <p className="text-xs font-medium">
              {nombre}
            </p>

            <div className="mt-2 space-y-1.5">

              <div className="flex items-center justify-between text-[10px]">
                <span className="text-muted-foreground">
                  REM
                </span>

                <span
                  className={
                    tieneRem
                      ? 'font-semibold text-primary'
                      : 'text-muted-foreground'
                  }
                >
                  {tieneRem ? '✓' : '—'}
                </span>
              </div>

              <div className="flex items-center justify-between text-[10px]">
                <span className="text-muted-foreground">
                  DÍAS
                </span>

                <span
                  className={
                    tieneDias
                      ? 'font-semibold text-primary'
                      : 'text-muted-foreground'
                  }
                >
                  {tieneDias ? '✓' : '—'}
                </span>
              </div>

            </div>
          </div>
        )
      })}
    </div>

    <div className="mt-4 flex flex-wrap gap-2">

      <div className="rounded-lg border border-border/60 bg-background px-3 py-2">
        <span className="text-[11px] text-muted-foreground">
          REM
        </span>
        <span className="ml-2 text-xs font-semibold">
          {trabajador.mesesRemuneraciones?.length ?? 0} meses
        </span>
      </div>

      <div className="rounded-lg border border-border/60 bg-background px-3 py-2">
        <span className="text-[11px] text-muted-foreground">
          DÍAS
        </span>
        <span className="ml-2 text-xs font-semibold">
          {trabajador.mesesDias?.length ?? 0} meses
        </span>
      </div>

      <div className="rounded-lg border border-border/60 bg-background px-3 py-2">
        <span className="text-[11px] text-muted-foreground">
          CONSOLIDADO
        </span>
        <span className="ml-2 text-xs font-semibold text-primary">
          {trabajador.mesesPresentes?.length ?? 0} meses
        </span>
      </div>

    </div>
  </div>
)}
    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-6 xl:grid-cols-6">
      {[
        ['Enero', trabajador.diasTrabajados.enero],
        ['Febrero', trabajador.diasTrabajados.febrero],
        ['Marzo', trabajador.diasTrabajados.marzo],
        ['Abril', trabajador.diasTrabajados.abril],
        ['Mayo', trabajador.diasTrabajados.mayo],
        ['Junio', trabajador.diasTrabajados.junio],
        ['Julio', trabajador.diasTrabajados.julio],
        ['Agosto', trabajador.diasTrabajados.agosto],
        ['Septiembre', trabajador.diasTrabajados.septiembre],
        ['Octubre', trabajador.diasTrabajados.octubre],
        ['Noviembre', trabajador.diasTrabajados.noviembre],
        ['Diciembre', trabajador.diasTrabajados.diciembre],
      ].map(([mes, dias]) => (
        <div
          key={mes}
          className="min-w-0 rounded-lg border border-border/60 bg-background p-3 text-center transition-colors hover:border-primary/30 hover:bg-primary/5"
        >
          <p className="text-[11px] font-medium text-muted-foreground">
            {mes}
          </p>
          <p className="mt-1 text-lg font-semibold">
            {dias}
          </p>
          <p className="text-[10px] text-muted-foreground">
            días
          </p>
        </div>
      ))}
    </div>
    <div className="mt-3 flex items-center justify-between rounded-lg border border-border/60 bg-background px-4 py-3">
      <span className="text-xs font-medium text-muted-foreground">
        Total anual registrado
      </span>
      <span className="text-sm font-bold text-primary">
        {trabajador.diasTrabajados.total} días
      </span>
    </div>
    {trabajador.incidencias && trabajador.incidencias.length > 0 && (
      <div className="mt-4 border-t border-border/60 pt-4">
        <p className="text-sm font-semibold">Incidencias</p>
        <p className="mt-1 text-xs text-muted-foreground">
          Ausencias y licencias registradas en el archivo importado
        </p>
        <div className="mt-3 space-y-2">
          {trabajador.incidencias.map((incidencia, incidenciaIndex) => (
            <div
              key={`${incidencia.codigo}-${incidenciaIndex}`}
              className="grid gap-2 rounded-lg border border-border/60 bg-background p-3 text-sm sm:grid-cols-[minmax(0,1fr)_auto_auto] sm:items-center"
            >
              <span className="min-w-0 font-medium">
                {incidencia.descripcion || incidencia.codigo || 'Incidencia registrada'}
              </span>
              <span className="whitespace-nowrap text-muted-foreground">
                {formatExcelDate(incidencia.fechaInicio)} - {formatExcelDate(incidencia.fechaFin)}
              </span>
              <span className="whitespace-nowrap font-semibold text-primary">
                {incidencia.cantidadDias} días
              </span>
            </div>
          ))}
        </div>
      </div>
    )}
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
  const [trabajadoresDias, setTrabajadoresDias] = useState<ImportedWorker[]>([])
  ; return (
  <div className="flex flex-col gap-7">
    <SectionTitle eyebrow="Cálculo · Paso 2 de 3" title="Distribución de utilidades" description="Revisa los parámetros, valida los factores de distribución y genera el cálculo anual del ejercicio." action={<div className="flex gap-2 sm:hidden">
    <Button variant="outline" size="sm">
    <Check className="mr-2 size-4" />Guardar</Button>
  <Button size="sm" onClick={onRun}>
    <Play className="mr-2 size-4" />Ejecutar</Button></div>} />
    {/* Flujo del cálculo */}
<div className="rounded-xl border border-border/60 bg-card p-5">
  <div className="mb-5 flex items-center justify-between">
    <div>
      <p className="text-sm font-semibold">Flujo del cálculo</p>
      <p className="mt-1 text-xs text-muted-foreground">
        Proceso de cálculo de utilidades
      </p>
    </div>
    <span className="text-xs font-medium text-primary">
      Paso 1 de 5
    </span>
  </div>
  <div className="flex items-center">
    {[
      "Carga",
      "Validación",
      "Días",
      "Remuneraciones",
      "Distribución",
    ].map((label, index) => (
      <div key={label} className="flex flex-1 items-center">
        <div className="flex min-w-0 flex-col items-center">
          <div
            className={`flex size-9 items-center justify-center rounded-full border text-sm font-semibold ${
              index === 0
                ? "border-primary bg-primary text-primary-foreground"
                : "border-border bg-muted/30 text-muted-foreground"
            }`}
          >
            {index + 1}
          </div>
          <span
            className={`mt-2 text-center text-xs font-medium ${
              index === 0
                ? "text-primary"
                : "text-muted-foreground"
            }`}
          >
            {label}
          </span>
        </div>
        {index < 4 && (
          <div className="mx-2 h-px flex-1 bg-border" />
        )}
      </div>
    ))}
  </div>
</div>
  <Card className="border-primary/30 bg-primary/5">
    <CardContent className="flex flex-col gap-4 p-5 sm:flex-row sm:items-center">
    <div className="flex size-11 items-center justify-center rounded-xl bg-primary/15 text-primary"><Upload /></div>
  <div className="flex-1">
    <p className="font-semibold">Base de trabajadores lista para procesar</p>
  <p className="mt-1 text-sm text-muted-foreground">Sin importaciones registradas</p></div>
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
    trabajadores={trabajadoresDias}
    setTrabajadores={setTrabajadoresDias}
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
          {errors.length === 0 && (
            <Badge
              variant="secondary"
              className="bg-emerald-500/10 text-emerald-400"
            >
              Validado
            </Badge>
          )}
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
                {totals.fund > 0 ? ((totals.distributed / totals.fund) * 100).toFixed(1) : '0.0'}%
              </span>
            </div>
            <Progress
              value={totals.fund > 0 ? (totals.distributed / totals.fund) * 100 : 0}
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