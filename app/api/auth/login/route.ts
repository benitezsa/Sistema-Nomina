import { NextResponse } from 'next/server'
import bcrypt from 'bcryptjs'
import type { RowDataPacket } from 'mysql2'
import pool from '@/lib/db'

type UsuarioDb = RowDataPacket & {
  id: number
  nombre: string
  apellido: string
  email: string
  rol: string
  password_hash: string
}

export async function POST(request: Request) {
  try {
    const body = await request.json().catch(() => ({}))
    const email = typeof body.email === 'string' ? body.email.trim().toLowerCase() : ''
    const password = typeof body.password === 'string' ? body.password : ''

    if (!email || !password) {
      return NextResponse.json(
        { success: false, message: 'Ingresa tu correo y contraseña.' },
        { status: 400 }
      )
    }

    const [rows] = await pool.query<UsuarioDb[]>(
      'SELECT id, nombre, apellido, email, rol, password_hash FROM usuarios WHERE email = ? AND activo = 1 LIMIT 1',
      [email]
    )
    const usuario = rows[0]

    if (!usuario || !(await bcrypt.compare(password, usuario.password_hash))) {
      return NextResponse.json(
        { success: false, message: 'Credenciales inválidas.' },
        { status: 401 }
      )
    }

    return NextResponse.json({
      success: true,
      user: {
        id: usuario.id,
        nombre: usuario.nombre,
        apellido: usuario.apellido,
        email: usuario.email,
        rol: usuario.rol,
      },
    })
  } catch (error) {
    console.error('Error en login:', error)
    return NextResponse.json(
      { success: false, message: 'Error interno al iniciar sesión.' },
      { status: 500 }
    )
  }
}