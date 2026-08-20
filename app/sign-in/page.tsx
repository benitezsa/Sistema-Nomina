import Link from 'next/link'
import { AuthForm } from '../../components/auth-form'

export default function SignInPage() {
  return (
    <main className="min-h-screen bg-background px-6 py-12 text-foreground">
      <section className="mx-auto flex min-h-[calc(100vh-6rem)] max-w-2xl flex-col justify-center">
        <div className="mb-10">
          <p className="mb-5 font-mono text-xs font-semibold uppercase tracking-[0.22em] text-primary">
            NóminaPro · Utilidades
          </p>

          <h1 className="text-balance text-4xl font-semibold tracking-tight">
            Bienvenido de nuevo
          </h1>

          <p className="mt-3 leading-6 text-muted-foreground">
            Accede a tu cuenta para continuar.
          </p>
        </div>

        <AuthForm />

        <div className="mt-6 flex items-center justify-between gap-4 text-sm">
          <Link
            href="/forgot-password"
            className="text-muted-foreground underline-offset-4 hover:text-foreground hover:underline"
          >
            ¿Olvidaste tu contraseña?
          </Link>

          <span className="text-muted-foreground">
            Acceso seguro
          </span>
        </div>
      </section>
    </main>
  )
}