import { describe, expect, it } from 'vitest'
import type { Employee } from '@/components/utilities/types'
import type { ImportedWorker } from '@/lib/import-utilidades'
import {
  calcularDias,
  calcularEmpleados,
  calcularRemuneracionComputable,
  calcularTrabajador,
  configuracion,
  normalizarTrabajador,
} from '@/lib/calculo'
import { calcularQuintaCategoria, calculateResults } from '@/lib/utilities-calculation'
import { mensual } from './fixtures'

function trabajador(overrides: Partial<ImportedWorker> = {}): ImportedWorker {
  return {
    dni: '12345678',
    apellidoPaterno: 'Perez',
    apellidoMaterno: 'Gomez',
    nombres: 'Juan',
    fechaInicio: '',
    fechaCese: '',
    remuneraciones: mensual({}),
    ...overrides,
  }
}

const config = configuracion({ anioEjercicio: 2025 })

describe('cálculo de días desde datos base', () => {
  it('recalcula el calendario laborable desde las fechas e ignora los totales del Excel', () => {
    const t = normalizarTrabajador(
      trabajador({
        fechaInicio: '2025-01-01',
        // Resultados del Excel: no deben alimentar el cálculo.
        diasPosibles: 999,
        diasEfectivosReferencia: 999,
        diasTrabajados: mensual({ enero: 999 }),
      }),
      config,
    )
    const dias = calcularDias(t, config)
    expect(dias.origen).toBe('calendario')
    expect(dias.laborables).toBe(261)
    expect(dias.laborables).not.toBe(999)
  })

  it('usa los días por mes del archivo solo cuando no hay fechas', () => {
    const t = normalizarTrabajador(
      trabajador({ laborablesPorMes: mensual({ enero: 20, febrero: 20 }) }),
      config,
    )
    const dias = calcularDias(t, config)
    expect(dias.origen).toBe('base-mensual')
    expect(dias.laborables).toBe(40)
  })

  it('descuenta los días no laborados de las incidencias', () => {
    const t = normalizarTrabajador(
      trabajador({
        fechaInicio: '2025-01-01',
        incidencias: [
          {
            dni: '12345678',
            codigo: '',
            descripcion: 'Faltas',
            cantidadDias: 3,
            fechaInicio: new Date('2025-01-06T00:00:00.000Z'),
            fechaFin: new Date('2025-01-08T00:00:00.000Z'),
            estado: 'identificada',
          },
        ],
      }),
      config,
    )
    const dias = calcularDias(t, config)
    expect(dias.noLaborados).toBe(3)
    expect(dias.efectivos).toBe(261 - 3)
  })
})

describe('remuneración computable desde conceptos', () => {
  it('suma los importes mensuales y deriva la mensual', () => {
    const t = normalizarTrabajador(
      trabajador({ remuneraciones: mensual({ enero: 1500, febrero: 1500, marzo: 3000 }) }),
      config,
    )
    const remuneracion = calcularRemuneracionComputable(t)
    expect(remuneracion.anual).toBe(6000)
    expect(remuneracion.mesesConRemuneracion).toBe(3)
    expect(remuneracion.mensual).toBe(2000)
  })

  it('conserva el total declarado cuando el archivo no trae desglose mensual', () => {
    const t = normalizarTrabajador(
      trabajador({ remuneraciones: { ...mensual({}), total: 24000 } }),
      config,
    )
    const remuneracion = calcularRemuneracionComputable(t)
    expect(remuneracion.anual).toBe(24000)
    expect(remuneracion.mensual).toBe(2000)
  })
})

describe('quinta categoría progresiva por tramos', () => {
  it('grava la utilidad con la tasa del tramo en el que cae', () => {
    const resultado = calcularQuintaCategoria(1000, 7 * 5350 + 10000)
    expect(resultado.impuesto).toBe(80)
    expect(resultado.tramosAplicados).toHaveLength(1)
    expect(resultado.tramosAplicados[0].tramo.tasa).toBe(0.08)
  })

  it('aplica cada tramo sobre su porción real', () => {
    const resultado = calcularQuintaCategoria(2000, 7 * 5350 + 26000)
    expect(resultado.impuesto).toBe(235)
    expect(resultado.tramosAplicados.map((tramo) => tramo.tramo.tasa)).toEqual([0.08, 0.14])
  })
})

describe('tope de la utilidad', () => {
  const empleado = (id: string): Employee => ({
    id,
    code: id,
    name: `Trabajador ${id}`,
    role: '',
    department: '',
    days: 261,
    diasEfectivos: 261,
    remuneration: 120000,
    remuneracionMensual: 10000,
    mesesConRemuneracion: 12,
    status: 'Completo',
  })

  it('usa la remuneración mensual, no la anual, para el tope de 18 remuneraciones', () => {
    const resultados = calculateResults([empleado('1'), empleado('2')], {
      income: 10_000_000,
      legalPercent: 10,
      averageEmployees: 100,
      daysBase: 360,
      remunerationMonths: 12,
      capMonths: 18,
    })
    for (const resultado of resultados) {
      expect(resultado.cap).toBe(180000)
      expect(resultado.finalAmount).toBe(180000)
      expect(resultado.capApplied).toBe(true)
    }
  })
})

describe('pipeline de cálculo', () => {
  it('produce empleados calculados sin leer totales precalculados', () => {
    const consolidado = [
      trabajador({
        fechaInicio: '2025-01-01',
        diasPosibles: 999,
        remuneraciones: mensual({ enero: 1000, febrero: 1000 }),
      }),
    ]
    const { empleados, calculados } = calcularEmpleados(consolidado, config)
    expect(calculados[0].dias.laborables).toBe(261)
    expect(empleados[0].diasLaborables).toBe(261)
    expect(empleados[0].remuneracionMensual).toBe(1000)
    expect(empleados[0].remuneration).toBe(2000)
  })

  it('proyecta el trabajador calculado con su jornada y estado', () => {
    const calculado = calcularTrabajador(
      normalizarTrabajador(trabajador({ fechaInicio: '2025-01-01', jornada: 6 }), config),
      config,
    )
    expect(calculado.dias.origen).toBe('calendario')
    expect(calculado.normalizado.jornada).toBe(6)
  })
})
