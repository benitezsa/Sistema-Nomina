import { NextResponse } from 'next/server'

import pool from '@/lib/db'

export const dynamic = 'force-dynamic'

function validarId(value: unknown): number | null {
  const numero = Number(value)
  return Number.isInteger(numero) && numero > 0 ? numero : null
}

export async function GET(request: Request) {
  const { searchParams } = new URL(request.url)
  const empresaId = validarId(searchParams.get('empresa_id'))
  const anio = validarId(searchParams.get('anio'))

  if (!empresaId || !anio) {
    return NextResponse.json(
      { error: 'empresa_id y anio son obligatorios' },
      { status: 400 },
    )
  }

  try {
    const [rows] = await pool.query(
      'SELECT id, empresa_id, anio, renta_neta, porcentaje_distribucion, fondo_total, fondo_dias, fondo_remuneraciones FROM ejercicios_utilidades WHERE empresa_id = ? AND anio = ? LIMIT 1',
      [empresaId, anio],
    ) as [Array<Record<string, unknown>>, unknown]

    if (rows.length === 0) {
      return NextResponse.json({ existe: false, ejercicio: null })
    }

    return NextResponse.json({ existe: true, ejercicio: rows[0] })
  } catch (error) {
    console.error('Error al consultar el ejercicio:', error)
    return NextResponse.json(
      { error: 'Error al consultar el ejercicio' },
      { status: 500 },
    )
  }
}

export async function POST(request: Request) {
  const body = await request.json().catch(() => null)
  const empresaId = validarId(body?.empresa_id)
  const anio = validarId(body?.anio)
  const rentaNeta = Number(body?.renta_neta)
  const porcentaje = Number(body?.porcentaje_distribucion)

  if (!empresaId || !anio) {
    return NextResponse.json(
      { error: 'empresa_id y anio son obligatorios' },
      { status: 400 },
    )
  }

  if (!Number.isFinite(rentaNeta) || rentaNeta < 0) {
    return NextResponse.json(
      { error: 'renta_neta debe ser un monto válido' },
      { status: 400 },
    )
  }

  if (!Number.isFinite(porcentaje) || porcentaje < 0) {
    return NextResponse.json(
      { error: 'porcentaje_distribucion debe ser un valor válido' },
      { status: 400 },
    )
  }

  const fondoTotal = rentaNeta * (porcentaje / 100)
  const fondoDias = fondoTotal * 0.5
  const fondoRemuneraciones = fondoTotal * 0.5

  try {
    const [result] = await pool.query(
      `INSERT INTO ejercicios_utilidades
        (empresa_id, anio, renta_neta, porcentaje_distribucion, fondo_total, fondo_dias, fondo_remuneraciones)
       VALUES (?, ?, ?, ?, ?, ?, ?)
       ON DUPLICATE KEY UPDATE
        renta_neta = VALUES(renta_neta),
        porcentaje_distribucion = VALUES(porcentaje_distribucion),
        fondo_total = VALUES(fondo_total),
        fondo_dias = VALUES(fondo_dias),
        fondo_remuneraciones = VALUES(fondo_remuneraciones)`,
      [
        empresaId,
        anio,
        Math.round(rentaNeta * 100) / 100,
        Math.round(porcentaje * 100) / 100,
        Math.round(fondoTotal * 100) / 100,
        Math.round(fondoDias * 100) / 100,
        Math.round(fondoRemuneraciones * 100) / 100,
      ],
    ) as [{ insertId: number }, unknown]

    return NextResponse.json({
      success: true,
      ejercicioId: result.insertId,
      fondo_total: Math.round(fondoTotal * 100) / 100,
      fondo_dias: Math.round(fondoDias * 100) / 100,
      fondo_remuneraciones: Math.round(fondoRemuneraciones * 100) / 100,
    })
  } catch (error) {
    console.error('Error al guardar el ejercicio:', error)
    return NextResponse.json(
      { error: 'Error al guardar el ejercicio' },
      { status: 500 },
    )
  }
}