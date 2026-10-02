import { describe, expect, it } from 'vitest'

import {
  consolidarTrabajadores,
  extraerDias,
  extraerRemuneraciones,
  formatearPeriodo,
  inferirAnioEjercicio,
  leerPeriodo,
  mesesTrabajados,
  vistaPorFuente,
  type ImportedWorker,
} from '@/lib/import-utilidades'
import { calcularDiasNoLaborados } from '@/lib/utilities-calculation'
import {
  libro,
  libroDias,
  libroDiasYFaltas,
  libroRemuneraciones,
  mensual,
  months,
  MESES_EN_CERO,
} from './fixtures'

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

describe('Meses trabajados: una sola fuente de verdad', () => {
  it('Caso A: un mes cuenta como trabajado solo si tiene días registrados', () => {
    const wb = libroDias([{ dni: '10000001', nombres: 'Ana Perez', dias: { enero: 23, febrero: 19, marzo: 21, abril: 22, mayo: 20, junio: 21, julio: 23, agosto: 20, septiembre: 22, octubre: 23, noviembre: 20, diciembre: 23 } }])
    const { trabajadores } = extraerDias(wb, [])

    expect(mesesTrabajados(buscar(trabajadores, '10000001'))).toHaveLength(12)
  })

  it('Caso B: los meses con 0 días no cuentan como trabajados', () => {
    const wb = libroDias([{ dni: '10000002', nombres: 'Luis Lopez', dias: { marzo: 21, abril: 22, mayo: 21, junio: 21, julio: 23, agosto: 21, septiembre: 20, octubre: 23, noviembre: 20, diciembre: 22 } }])
    const { trabajadores } = extraerDias(wb, [])

    const trabajador = buscar(trabajadores, '10000002')
    const trabajados = mesesTrabajados(trabajador)

    expect(trabajados).toHaveLength(10)
    expect(trabajados).not.toContain('enero')
    expect(trabajados).not.toContain('febrero')
    expect(trabajados[0]).toBe('marzo')
  })

  it('Caso B: tener una fila en el Excel no implica haber trabajado', () => {
    // El archivo trae los 12 meses, pero dos están en cero.
    const sinFaltas = {
      diasTrabajados: mensual({ enero: 0, febrero: 0, marzo: 20, abril: 20, mayo: 20, junio: 20, julio: 20, agosto: 20, septiembre: 20, octubre: 20, noviembre: 20, diciembre: 20 }),
      mesesPresentes: months as unknown as string[],
    }

    expect(mesesTrabajados(sinFaltas)).toHaveLength(10)
  })

  it('Caso G: distintas cantidades de meses trabajados', () => {
    const wb = libroDias([
      { dni: '30000001', nombres: 'Uno Uno', dias: Object.fromEntries(months.map((m) => [m, 20])) },
      { dni: '30000002', nombres: 'Dos Dos', dias: { junio: 20, julio: 20, agosto: 20 } },
      { dni: '30000003', nombres: 'Tres Tres', dias: { diciembre: 15 } },
    ])
    const { trabajadores } = extraerDias(wb, [])

    expect(mesesTrabajados(buscar(trabajadores, '30000001'))).toHaveLength(12)
    expect(mesesTrabajados(buscar(trabajadores, '30000002'))).toHaveLength(3)
    expect(mesesTrabajados(buscar(trabajadores, '30000003'))).toHaveLength(1)
  })

  it('sin desglose mensual cae a la permanencia que el motor consolidó', () => {
    expect(mesesTrabajados({ mesesDias: ['abril', 'mayo'] })).toEqual(['abril', 'mayo'])
    expect(mesesTrabajados({ mesesPresentes: ['octubre'] })).toEqual(['octubre'])
  })

  it('el override de la interfaz manda sobre los datos importados', () => {
    const trabajador: ImportedWorker = {
      dni: '40000001',
      apellidoPaterno: 'A',
      apellidoMaterno: 'B',
      nombres: 'C',
      fechaInicio: '',
      fechaCese: '',
      remuneraciones: { ...MESES_EN_CERO },
      diasTrabajados: mensual({ enero: 20, febrero: 20 }),
    }

    expect(mesesTrabajados(trabajador, ['enero'])).toEqual(['enero'])
  })
})

describe('Incidencias: el período del archivo se conserva', () => {
  it('Caso C: una matriz mensual conserva el mes y los días de cada incidencia', () => {
    const wb = libroDiasYFaltas([
      { dni: '50000001', nombres: 'Ana Perez', dias: { enero: 23, febrero: 19, marzo: 21, abril: 22, mayo: 20, junio: 21, julio: 23, agosto: 20, septiembre: 22, octubre: 23, noviembre: 20, diciembre: 23 }, faltas: { febrero: 1, mayo: 2, agosto: 1 } },
    ])
    const { trabajadores } = extraerDias(wb, [])

    const trabajador = buscar(trabajadores, '50000001')
    const incidencias = trabajador.incidencias ?? []
    const periodos = incidencias.map((inc) => formatearPeriodo(inc.periodo)).sort()

    expect(incidencias).toHaveLength(3)
    expect(periodos).toEqual(['Agosto', 'Febrero', 'Mayo'])
    expect(incidencias.map((inc) => inc.cantidadDias).sort((a, b) => a - b)).toEqual([1, 1, 2])
    // No se inventan fechas de inicio ni de fin.
    expect(incidencias.every((inc) => !inc.fechaInicio && !inc.fechaFin)).toBe(true)
    expect(diasNoLaborados(trabajador)).toBe(4)
    expect(trabajador.diasTrabajados?.total).toBe(257)
    expect(mesesTrabajados(trabajador)).toHaveLength(12)
  })

  it('Caso C: el año del período se toma del archivo, no se inventa', () => {
    const sinFechas = libro([{ nombre: 'DIAS', encabezados: ['DNI', 'Nombres'], filas: [['50000003', 'Ana Perez']] }])
    expect(inferirAnioEjercicio(sinFechas)).toBeNull()

    const conFechas = libroRemuneraciones([{ dni: '50000002', nombres: 'Luis Lopez', ingreso: '2023-04-01', remuneraciones: { abril: 100 } }])
    expect(inferirAnioEjercicio(conFechas)).toBe(2023)
  })

  it('Caso D: una tabla PERIODO + TIPO + DIAS conserva período y días', () => {
    const wb = libro([{
      nombre: 'INCIDENCIAS',
      encabezados: ['DNI', 'Nombres', 'PERIODO', 'CONCEPTO', 'NUM. DIAS'],
      filas: [
        ['60000001', 'Ana Perez', '2025-02', 'FALTA', 1],
        ['60000001', 'Ana Perez', '2025-05', 'FALTA', 2],
        ['60000001', 'Ana Perez', '2025-08', 'FALTA', 1],
        ['60000002', 'Luis Lopez', '2025-03', 'VACACIONES', 5],
      ],
    }])
    const { trabajadores } = extraerDias(wb, [])

    const ana = buscar(trabajadores, '60000001')
    const incidencias = ana.incidencias ?? []
    const periodos = incidencias.map((inc) => formatearPeriodo(inc.periodo)).sort()

    expect(incidencias).toHaveLength(3)
    expect(periodos).toEqual(['Agosto 2025', 'Febrero 2025', 'Mayo 2025'])
    expect(diasNoLaborados(ana)).toBe(4)
    expect(incidencias.every((inc) => inc.estado === 'identificada')).toBe(true)
  })

  it('Caso D: las fechas del archivo se conservan y tienen prioridad sobre el período', () => {
    const wb = libro([{
      nombre: 'INCIDENCIAS',
      encabezados: ['DNI', 'Nombres', 'PERIODO', 'CONCEPTO', 'NUM. DIAS', 'F. INICIO', 'F. FIN'],
      filas: [
        ['70000001', 'Ana Perez', '2025-05', 'LICENCIA MEDICA', 3, '2025-05-05', '2025-05-07'],
      ],
    }])
    const { trabajadores } = extraerDias(wb, [])

    const incidencia = (buscar(trabajadores, '70000001').incidencias ?? [])[0]

    expect(formatearPeriodo(incidencia.periodo)).toBe('Mayo 2025')
    expect(incidencia.tipo).toBe('licencia_medica')
  })

  it('sin período ni fechas la incidencia queda como pendiente, no se inventa', () => {
    const wb = libro([{
      nombre: 'INCIDENCIAS',
      encabezados: ['DNI', 'Nombres', 'CONCEPTO', 'NUM. DIAS'],
      filas: [
        ['80000001', 'Ana Perez', 'FALTA', 2],
        ['80000001', 'Ana Perez', 'FALTA', ''],
      ],
    }])
    const { trabajadores } = extraerDias(wb, [])

    const incidencias = buscar(trabajadores, '80000001').incidencias ?? []
    const identificada = incidencias.find((inc) => inc.cantidadDias > 0)
    const pendiente = incidencias.find((inc) => inc.cantidadDias === 0)

    expect(incidencias).toHaveLength(2)
    expect(identificada?.estado).toBe('identificada')
    expect(identificada?.periodo ?? null).toBeNull()
    // Sin días, sin rango de fechas y sin período: no se inventa nada.
    expect(pendiente?.estado).toBe('no_identificada')
    expect(pendiente?.motivo).toBeTruthy()
    expect(diasNoLaborados(buscar(trabajadores, '80000001'))).toBe(2)
  })

  it('los días no laborados siguen la fórmula días registrados menos días no laborados', () => {
    const wb = libroDiasYFaltas([
      { dni: '90000001', nombres: 'Ana Perez', dias: { enero: 23, febrero: 19, marzo: 21 }, faltas: { febrero: 1, mayo: 2, agosto: 1 } },
    ])
    const { trabajadores } = extraerDias(wb, [])
    const trabajador = buscar(trabajadores, '90000001')
    const posibles = trabajador.diasTrabajados?.total ?? 0

    expect(posibles).toBe(63)
    expect(diasNoLaborados(trabajador)).toBe(4)
    expect(Math.max(0, posibles - diasNoLaborados(trabajador))).toBe(59)
  })
})

describe('Incidencias: lectura de período', () => {
  it('interpreta año y mes en distintos formatos', () => {
    expect(leerPeriodo('2025-02')).toEqual({ texto: '2025-02', anio: 2025, mes: 2 })
    expect(leerPeriodo('02/2025')).toEqual({ texto: '02/2025', anio: 2025, mes: 2 })
    expect(leerPeriodo('FEBRERO')).toEqual({ texto: 'FEBRERO', anio: null, mes: 2 })
    expect(leerPeriodo('FEBRERO 2025')).toEqual({ texto: 'FEBRERO 2025', anio: 2025, mes: 2 })
    expect(leerPeriodo('SETIEMBRE')).toEqual({ texto: 'SETIEMBRE', anio: null, mes: 9 })
  })

  it('no inventa el año si el archivo no lo aporta', () => {
    const periodo = leerPeriodo('FEBRERO')

    expect(formatearPeriodo(periodo)).toBe('Febrero')
    expect(formatearPeriodo(periodo, 2025)).toBe('Febrero 2025')
  })

  it('conserva un texto de período que no reconoce', () => {
    const periodo = leerPeriodo('PRIMERA QUINCENA')

    expect(periodo?.texto).toBe('PRIMERA QUINCENA')
    expect(periodo?.mes).toBeNull()
    expect(formatearPeriodo(periodo)).toBe('PRIMERA QUINCENA')
  })

  it('no toma un número suelto por un año', () => {
    expect(leerPeriodo('30')?.anio).toBeNull()
    expect(leerPeriodo('SEMANA 12')?.anio).toBeNull()
    // Junto a un nombre de mes, el año de dos dígitos sí es un año.
    expect(leerPeriodo('FEB-25')).toEqual({ texto: 'FEB-25', anio: 2025, mes: 2 })
  })

  it('una fecha completa en la columna de período corresponde a su mes', () => {
    expect(leerPeriodo('15/02/2025')).toEqual({ texto: '15/02/2025', anio: 2025, mes: 2 })
  })

  it('sin período devuelve vacío y no inventa', () => {
    expect(leerPeriodo(null)).toBeNull()
    expect(leerPeriodo('')).toBeNull()
    expect(formatearPeriodo(null)).toBe('')
  })
})

describe('Fecha de ingreso: fuente única', () => {
  it('Caso E: un trabajador con fecha de ingreso la conserva', () => {
    const rem = libroRemuneraciones([{ dni: '11000001', nombres: 'Ana Perez', ingreso: '2024-01-15', remuneraciones: { enero: 1000 } }])
    const remuneracion = extraerRemuneraciones(rem, []).trabajadores
    const conDias = extraerDias(rem, remuneracion).trabajadores

    expect(buscar(conDias, '11000001').fechaInicio).toBe('2024-01-15')
  })

  it('Caso F: un trabajador sin fecha de ingreso no inventa una', () => {
    const wb = libroDias([{ dni: '11000002', nombres: 'Luis Lopez', dias: { enero: 20 } }])
    const { trabajadores } = extraerDias(wb, [])

    const trabajador = buscar(trabajadores, '11000002')
    const ingreso = trabajador.fechaInicio

    expect(ingreso === '' || ingreso === null || ingreso === undefined).toBe(true)
  })

  it('una carga de remuneraciones sin columna de fechas no borra el dato previo', () => {
    const remConFecha = libroRemuneraciones([
      { dni: '11000003', nombres: 'Ana Perez', ingreso: '2024-03-01', remuneraciones: { enero: 1000 } },
    ])
    const previas = extraerRemuneraciones(remConFecha, []).trabajadores
    expect(buscar(previas, '11000003').fechaInicio).toBe('2024-03-01')

    const remSinFecha = libroRemuneraciones([
      { dni: '11000003', nombres: 'Ana Perez', remuneraciones: { febrero: 2000 } },
    ])
    const despues = extraerRemuneraciones(remSinFecha, previas).trabajadores

    expect(buscar(despues, '11000003').fechaInicio).toBe('2024-03-01')
  })

  it('la carga de días sin fechas hereda las de remuneraciones en la vista consolidada', () => {
    const rem = libroRemuneraciones([
      { dni: '11000004', nombres: 'Ana Perez', ingreso: '2024-05-10', remuneraciones: { enero: 1000 } },
    ])
    const remuneraciones = extraerRemuneraciones(rem, []).trabajadores

    // El archivo de días no trae columna de fecha de ingreso.
    const wbDias = libroDias([{ dni: '11000004', nombres: 'Ana Perez', dias: { enero: 20 } }])
    const dias = extraerDias(wbDias, []).trabajadores
    expect(buscar(dias, '11000004').fechaInicio).toBe('')

    // La tarjeta de días muestra la vista consolidada: ya no dice "No disponible".
    const vista = vistaPorFuente(consolidarTrabajadores(dias, remuneraciones), dias)
    expect(buscar(vista, '11000004').fechaInicio).toBe('2024-05-10')
  })

  it('la vista de una carga solo lista a los trabajadores que esa carga trajo', () => {
    const rem = libroRemuneraciones([
      { dni: '11000005', nombres: 'Ana Perez', ingreso: '2024-01-05', remuneraciones: { enero: 1000 } },
      { dni: '11000006', nombres: 'Luis Lopez', remuneraciones: { enero: 1000 } },
    ])
    const wbDias = libroDias([{ dni: '11000005', nombres: 'Ana Perez', dias: { enero: 20 } }])

    const remuneraciones = extraerRemuneraciones(rem, []).trabajadores
    const dias = extraerDias(wbDias, []).trabajadores
    const consolidada = consolidarTrabajadores(dias, remuneraciones)

    expect(consolidada).toHaveLength(2)
    expect(vistaPorFuente(consolidada, dias).map((t) => t.dni)).toEqual(['11000005'])
    expect(vistaPorFuente(consolidada, remuneraciones).map((t) => t.dni)).toEqual([
      '11000005',
      '11000006',
    ])
  })
})

describe('Independencia entre remuneraciones, días e incidencias', () => {
  it('Caso H: remuneración presente pero sin incidencia', () => {
    const rem = libroRemuneraciones([{ dni: '12000001', nombres: 'Ana Perez', remuneraciones: { enero: 1000, febrero: 1000 } }])
    const remuneraciones = extraerRemuneraciones(rem, []).trabajadores
    const { trabajadores } = extraerDias(rem, remuneraciones)

    const trabajador = buscar(trabajadores, '12000001')

    expect(trabajador.remuneraciones.total).toBe(2000)
    expect(trabajador.incidencias ?? []).toHaveLength(0)
    // La remuneración no genera días ni incidencias.
    expect(trabajador.diasTrabajados ?? undefined).toBeUndefined()
  })

  it('Caso I: incidencia presente pero sin remuneración', () => {
    const wb = libroDiasYFaltas([
      { dni: '12000002', nombres: 'Luis Lopez', dias: { marzo: 21 }, faltas: { marzo: 1 } },
    ])
    const { trabajadores } = extraerDias(wb, [])

    const trabajador = buscar(trabajadores, '12000002')

    expect(trabajador.incidencias).toHaveLength(1)
    expect(trabajador.remuneraciones.total).toBe(0)
    // La incidencia no se convierte en remuneración.
    expect(months.every((mes) => trabajador.remuneraciones[mes] === 0)).toBe(true)
  })

  it('un concepto de remuneración no se convierte en incidencia', () => {
    const wb = libro([{
      nombre: 'INCIDENCIAS',
      encabezados: ['DNI', 'Nombres', 'PERIODO', 'CONCEPTO', 'NUM. DIAS'],
      filas: [['13000001', 'Ana Perez', '2025-02', 'SUELDO BASICO', 1000]],
    }])
    const { trabajadores } = extraerDias(wb, [])

    const incidencias = buscar(trabajadores, '13000001').incidencias ?? []

    expect(incidencias.every((inc) => inc.estado !== 'identificada')).toBe(true)
    expect(trabajadores.every((t) => t.remuneraciones.total === 0)).toBe(true)
  })

  it('los días registrados no se derivan de la remuneración', () => {
    const wb = libroDias([{ dni: '14000001', nombres: 'Ana Perez', dias: { enero: 23, febrero: 19 } }])
    const { trabajadores } = extraerDias(wb, [])

    const trabajador = buscar(trabajadores, '14000001')

    expect(trabajador.diasTrabajados?.total).toBe(42)
    expect(trabajador.remuneraciones.total).toBe(0)
  })
})
