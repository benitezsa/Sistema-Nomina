import { existsSync } from 'node:fs'
import path from 'node:path'
import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'

import {
  consolidarTrabajadores,
  extraerDias,
  extraerRemuneraciones,
  formatearPeriodo,
  mesesTrabajados,
  type ImportedWorker,
} from '@/lib/import-utilidades'
import { calcularDiasNoLaborados } from '@/lib/utilities-calculation'

const ARCHIVO_PRUEBA = path.resolve(process.cwd(), 'Excel_Prueba_Modulo_Utilidades_2025.xlsx')

function buscar(trabajadores: ImportedWorker[], dni: string): ImportedWorker {
  const trabajador = trabajadores.find((t) => t.dni === dni)
  if (!trabajador) throw new Error(`No se encontró al trabajador ${dni}`)
  return trabajador
}

function diasNoLaborados(trabajador: ImportedWorker): number {
  return calcularDiasNoLaborados(
    (trabajador.incidencias ?? []).map((inc) => ({
      cantidadDias: inc.cantidadDias,
      diasNeto: inc.diasNeto,
    })),
    5,
    new Set(trabajador.feriados ?? []),
  )
}

async function cargar() {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile(ARCHIVO_PRUEBA)
  const remuneraciones = extraerRemuneraciones(wb, [], { loteId: 'prueba' })
  const dias = extraerDias(wb, remuneraciones.trabajadores)
  return {
    remuneraciones: remuneraciones.trabajadores,
    dias: dias.trabajadores,
    // La vista de Validación consolida días y remuneraciones con el motor.
    consolidado: consolidarTrabajadores(dias.trabajadores, remuneraciones.trabajadores),
  }
}

describe.skipIf(!existsSync(ARCHIVO_PRUEBA))('Excel de prueba del módulo de Utilidades', () => {
  it('70000001: días, no laborados y efectivos se mantienen', async () => {
    const { consolidado } = await cargar()
    const trabajador = buscar(consolidado, '70000001')
    const posibles = trabajador.diasPosibles ?? trabajador.diasTrabajados?.total ?? 0
    const noLaborados = diasNoLaborados(trabajador)

    expect(trabajador.diasTrabajados?.total).toBe(257)
    expect(noLaborados).toBe(4)
    expect(Math.max(0, posibles - noLaborados)).toBe(253)
  })

  it('70000001: la fecha de ingreso es la misma en todas las secciones', async () => {
    const { remuneraciones, consolidado } = await cargar()
    const desdeRemuneraciones = buscar(remuneraciones, '70000001')
    const desdeValidacion = buscar(consolidado, '70000001')

    expect(desdeRemuneraciones.fechaInicio).toBe('2024-01-15')
    expect(desdeValidacion.fechaInicio).toBe('2024-01-15')
    // La carga de días no borra la fecha que aporta la de remuneraciones.
    expect(desdeRemuneraciones.fechaInicio).toBe(desdeValidacion.fechaInicio)
  })

  it('70000001: los doce meses tienen días registrados', async () => {
    const { consolidado } = await cargar()
    const trabajador = buscar(consolidado, '70000001')

    expect(mesesTrabajados(trabajador)).toHaveLength(12)
  })

  it('70000001: las incidencias conservan su período', async () => {
    const { consolidado } = await cargar()
    const incidencias = buscar(consolidado, '70000001').incidencias ?? []

    expect(incidencias.map((inc) => `${formatearPeriodo(inc.periodo)} · ${inc.cantidadDias}`).sort()).toEqual([
      'Agosto 2025 · 1',
      'Febrero 2025 · 1',
      'Mayo 2025 · 2',
    ])
  })

  it('70000004: días, no laborados y efectivos se mantienen', async () => {
    const { consolidado } = await cargar()
    const trabajador = buscar(consolidado, '70000004')
    const posibles = trabajador.diasPosibles ?? trabajador.diasTrabajados?.total ?? 0
    const noLaborados = diasNoLaborados(trabajador)

    expect(trabajador.diasTrabajados?.total).toBe(214)
    expect(noLaborados).toBe(4)
    expect(Math.max(0, posibles - noLaborados)).toBe(210)
  })

  it('70000004: la fecha de ingreso es la misma en todas las secciones', async () => {
    const { remuneraciones, consolidado } = await cargar()

    expect(buscar(remuneraciones, '70000004').fechaInicio).toBe('2025-03-01')
    expect(buscar(consolidado, '70000004').fechaInicio).toBe('2025-03-01')
  })

  it('70000004: diez meses trabajados y dos sin trabajar, sin variables paralelas', async () => {
    const { consolidado } = await cargar()
    const trabajador = buscar(consolidado, '70000004')
    const trabajados = mesesTrabajados(trabajador)

    expect(trabajados).toHaveLength(10)
    expect(trabajados).not.toContain('enero')
    expect(trabajados).not.toContain('febrero')
    // El resumen se calcula con los mismos meses que usa "Períodos trabajados".
    expect(mesesTrabajados(trabajador)).toEqual(trabajador.mesesPresentes ?? [])
  })

  it('70000004: las incidencias conservan su período', async () => {
    const { consolidado } = await cargar()
    const incidencias = buscar(consolidado, '70000004').incidencias ?? []

    expect(incidencias.map((inc) => `${formatearPeriodo(inc.periodo)} · ${inc.cantidadDias}`).sort()).toEqual([
      'Diciembre 2025 · 1',
      'Mayo 2025 · 1',
      'Septiembre 2025 · 2',
    ])
  })
})

describe('Consolidado de las cargas', () => {
  const SIN_DIAS = {
    diasTrabajados: undefined,
    incidencias: undefined,
  }

  it('la fecha de ingreso no se pierde cuando una carga la trae vacía', () => {
    const conIngreso: ImportedWorker = {
      dni: '1',
      apellidoPaterno: 'Perez',
      apellidoMaterno: 'Quispe',
      nombres: 'Ana',
      fechaInicio: '2024-01-15',
      fechaCese: '',
      remuneraciones: { enero: 1000, febrero: 1000, marzo: 0, abril: 0, mayo: 0, junio: 0, julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 2000 },
      mesesPresentes: ['enero', 'febrero'],
    }
    const sinFecha: ImportedWorker = {
      ...conIngreso,
      fechaInicio: '',
      remuneraciones: { enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0, julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 0 },
      ...SIN_DIAS,
    }

    const [consolidado] = consolidarTrabajadores([sinFecha], [conIngreso])

    expect(consolidado.fechaInicio).toBe('2024-01-15')
    expect(consolidado.remuneraciones.total).toBe(2000)
  })

  it('la carga de días va primero: sus meses trabajados mandan sobre los de remuneración', () => {
    const base: ImportedWorker = {
      dni: '1',
      apellidoPaterno: 'Condori',
      apellidoMaterno: 'Soto',
      nombres: 'Daniel',
      fechaInicio: '2025-03-01',
      fechaCese: '',
      remuneraciones: { enero: 2000, febrero: 2000, marzo: 2000, abril: 2000, mayo: 2000, junio: 2000, julio: 2000, agosto: 2000, septiembre: 2000, octubre: 2000, noviembre: 2000, diciembre: 2000, total: 24000 },
      mesesPresentes: ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
    }
    const conDias: ImportedWorker = {
      ...base,
      remuneraciones: { enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0, julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 0 },
      diasTrabajados: { enero: 0, febrero: 0, marzo: 21, abril: 22, mayo: 21, junio: 21, julio: 23, agosto: 21, septiembre: 20, octubre: 23, noviembre: 20, diciembre: 22, total: 214 },
      mesesDias: ['marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
      mesesPresentes: ['marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'],
    }

    const [consolidado] = consolidarTrabajadores([conDias], [base])

    expect(mesesTrabajados(consolidado)).toHaveLength(10)
    expect(consolidado.mesesPresentes).toEqual(mesesTrabajados(consolidado))
    expect(consolidado.fechaInicio).toBe('2025-03-01')
  })

  it('las incidencias de varias cargas se conservan y no se duplican entre sí', () => {
    const base: ImportedWorker = {
      dni: '1',
      apellidoPaterno: 'Quispe',
      apellidoMaterno: 'Mamani',
      nombres: 'Ana',
      fechaInicio: '',
      fechaCese: '',
      remuneraciones: { enero: 0, febrero: 0, marzo: 0, abril: 0, mayo: 0, junio: 0, julio: 0, agosto: 0, septiembre: 0, octubre: 0, noviembre: 0, diciembre: 0, total: 0 },
    }
    const conFaltas: ImportedWorker = {
      ...base,
      incidencias: [
        { dni: '1', codigo: '', descripcion: 'Faltas', cantidadDias: 1, fechaInicio: null, fechaFin: null, periodo: { texto: 'Febrero', anio: 2025, mes: 2 } },
      ],
    }
    const conVacaciones: ImportedWorker = {
      ...base,
      incidencias: [
        { dni: '1', codigo: '', descripcion: 'Vacaciones', cantidadDias: 3, fechaInicio: null, fechaFin: null, periodo: { texto: 'Mayo', anio: 2025, mes: 5 } },
      ],
    }

    const [consolidado] = consolidarTrabajadores([conFaltas], [conVacaciones])

    expect(consolidado.incidencias?.map((inc) => inc.descripcion)).toEqual(['Faltas', 'Vacaciones'])
    expect(calcularDiasNoLaborados(
      (consolidado.incidencias ?? []).map((inc) => ({ cantidadDias: inc.cantidadDias })),
    )).toBe(4)
  })
})
