'use client'

import { createContext, useCallback, useContext, useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'

export type SesionUsuario = {
  id: number
  nombre: string
  apellido: string
  email: string
  rol: string
}

const SESSION_KEY = 'utilidades-usuario'

type AuthContextValue = {
  user: SesionUsuario | null
  cargando: boolean
  login: (email: string, password: string) => Promise<SesionUsuario>
  logout: () => void
}

const AuthContext = createContext<AuthContextValue | null>(null)

function cargarSesion(): SesionUsuario | null {
  try {
    const raw = window.localStorage.getItem(SESSION_KEY)
    return raw ? (JSON.parse(raw) as SesionUsuario) : null
  } catch {
    return null
  }
}

export function AuthProvider({ children }: { children: React.ReactNode }) {
  const router = useRouter()
  const [user, setUser] = useState<SesionUsuario | null>(null)
  const [cargando, setCargando] = useState(true)

  useEffect(() => {
    setUser(cargarSesion())
    setCargando(false)
  }, [])

  const login = useCallback(async (email: string, password: string) => {
    const respuesta = await fetch('/api/auth/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password }),
    })
    const datos = await respuesta.json().catch(() => ({}))
    if (!respuesta.ok || !datos?.success) {
      throw new Error(datos?.message || 'No se pudo iniciar sesión.')
    }
    const usuario = datos.user as SesionUsuario
    window.localStorage.setItem(SESSION_KEY, JSON.stringify(usuario))
    setUser(usuario)
    return usuario
  }, [])

  const logout = useCallback(() => {
    window.localStorage.removeItem(SESSION_KEY)
    setUser(null)
    router.push('/sign-in')
  }, [router])

  return (
    <AuthContext.Provider value={{ user, cargando, login, logout }}>
      {children}
    </AuthContext.Provider>
  )
}

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext)
  if (!ctx) {
    throw new Error('useAuth debe usarse dentro de un AuthProvider')
  }
  return ctx
}