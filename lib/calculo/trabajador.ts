import type { Employee } from '@/components/utilities/types'
import type { ConfiguracionCalculo } from './config'
import { calcularDias, type ResultadoDias } from './dias'
import type { TrabajadorNormalizado } from './modelo'
import { calcularRemuneracionComputable, type ResultadoRemuneracion } from './remuneraciones'

export type TrabajadorCalculado = {
  normalizado: TrabajadorNormalizado
  dias: ResultadoDias
  remuneracion: ResultadoRemuneracion
}

export function calcularTrabajador(
  t: TrabajadorNormalizado,
  config: ConfiguracionCalculo,
): TrabajadorCalculado {
  return {
    normalizado: t,
    dias: calcularDias(t, config),
    remuneracion: calcularRemuneracionComputable(t),
  }
}

export function calcularTrabajadores(
  trabajadores: TrabajadorNormalizado[],
  config: ConfiguracionCalculo,
): TrabajadorCalculado[] {
  return trabajadores.map((t) => calcularTrabajador(t, config))
}

/** Proyecta el trabajador calculado al tipo que consume el motor de resultados. */
export function aEmployee(calculado: TrabajadorCalculado): Employee {
  const t = calculado.normalizado
  const nombre = `${t.apellidoPaterno} ${t.apellidoMaterno}`.trim() || t.nombres
  return {
    id: t.dni,
    code: t.dni,
    name: nombre,
    role: '',
    department: '',
    days: calculado.dias.efectivos,
    remuneration: calculado.remuneracion.anual,
    status:
      calculado.dias.laborables > 0 && calculado.remuneracion.anual > 0 ? 'Completo' : 'Pendiente',
    jornada: t.jornada,
    diasLaborables: calculado.dias.laborables,
    diasNoLaborados: calculado.dias.noLaborados,
    diasEfectivos: calculado.dias.efectivos,
    remuneracionMensual: calculado.remuneracion.mensual,
    mesesConRemuneracion: calculado.remuneracion.mesesConRemuneracion,
  }
}
