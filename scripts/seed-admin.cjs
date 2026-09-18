const fs = require('fs')
const path = require('path')
const mysql = require('mysql2/promise')
const bcrypt = require('bcryptjs')

function cargarEnv(ruta) {
  const env = {}
  const contenido = fs.readFileSync(ruta, 'utf8')
  for (const linea of contenido.split(/\r?\n/)) {
    const m = linea.match(/^\s*([\w.-]+)\s*=\s*(.*?)\s*$/)
    if (m) env[m[1]] = m[2].replace(/^['"]|['"]$/g, '')
  }
  return env
}

async function main() {
  const env = cargarEnv(path.join(__dirname, '..', '.env.local'))
  const email = (process.env.SEED_EMAIL || 'admin@sistema.com').toLowerCase()
  const password = process.env.SEED_PASSWORD || 'Admin123!'

  if (!email || !password) {
    throw new Error('Faltan variables SEED_EMAIL / SEED_PASSWORD')
  }

  const hash = await bcrypt.hash(password, 10)
  const c = await mysql.createConnection({
    host: env.DB_HOST,
    port: Number(env.DB_PORT || 3306),
    user: env.DB_USER,
    password: env.DB_PASSWORD,
    database: env.DB_NAME,
  })

  const [rows] = await c.query('SELECT id FROM usuarios WHERE email = ? LIMIT 1', [email])
  if (rows.length > 0) {
    await c.query('UPDATE usuarios SET password_hash = ?, activo = 1, rol = ? WHERE email = ?', [hash, 'admin', email])
    console.log('Usuario actualizado:', email)
  } else {
    await c.query(
      'INSERT INTO usuarios (nombre, apellido, email, password_hash, rol, activo) VALUES (?, ?, ?, ?, ?, 1)',
      ['Administrador', 'Sistema', email, hash, 'admin']
    )
    console.log('Usuario creado:', email)
  }

  await c.end()
  console.log('Credenciales de acceso ->', email, '/', password)
}

main().catch((e) => {
  console.error(e.message)
  process.exit(1)
})