'use client'

import { FormEvent, useState } from 'react'
import { useRouter } from 'next/navigation'

export function AuthForm() {
  const router = useRouter()

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    setError('')

    if (!email || !password) {
      setError('Ingresa tu correo y contraseña.')
      return
    }

    setLoading(true)

    // Inicio de sesión temporal.
    // La autenticación real se conectará posteriormente.
    setTimeout(() => {
      router.push('/')
      router.refresh()
    }, 500)
  }

  return (
    <form onSubmit={handleSubmit} className="grid gap-5">
      <div className="grid gap-2">
        <label
          htmlFor="email"
          className="text-sm font-medium text-foreground"
        >
          Correo electrónico
        </label>

        <input
          id="email"
          type="email"
          autoComplete="email"
          required
          value={email}
          onChange={(event) => setEmail(event.target.value)}
          placeholder="nombre@empresa.com"
          className="h-12 rounded-md border border-input bg-background px-4 text-base outline-none transition focus:border-primary focus:ring-2 focus:ring-ring/20"
        />
      </div>

      <div className="grid gap-2">
        <label
          htmlFor="password"
          className="text-sm font-medium text-foreground"
        >
          Contraseña
        </label>

        <input
          id="password"
          type="password"
          autoComplete="current-password"
          required
          value={password}
          onChange={(event) => setPassword(event.target.value)}
          placeholder="••••••••"
          className="h-12 rounded-md border border-input bg-background px-4 text-base outline-none transition focus:border-primary focus:ring-2 focus:ring-ring/20"
        />
      </div>

      {error && (
        <p className="text-sm text-destructive">
          {error}
        </p>
      )}

      <button
        type="submit"
        disabled={loading}
        className="h-12 rounded-md bg-primary px-5 font-semibold text-primary-foreground transition hover:bg-primary/90 disabled:cursor-not-allowed disabled:opacity-60"
      >
        {loading ? 'Ingresando…' : 'Iniciar sesión'}
      </button>
    </form>
  )
}