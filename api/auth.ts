import { createHmac, timingSafeEqual } from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'

export type UserRole = 'user' | 'manager' | 'super_user'
export type Session = { role: UserRole; userId: string }

function accessCodes() {
  try {
    const values = fs.readFileSync(path.resolve(process.cwd(), 'xmlfiles', 'admin-pwd.txt'), 'utf8')
      .split(/\r?\n/).map((value: string) => value.trim()).filter(Boolean)
    if (values.length >= 2) return { manager: values[0], superUser: values[1] }
  } catch {
    // Production provides the same values through environment variables.
  }
  return { manager: process.env.ADMIN_PASSWORD || '', superUser: process.env.SUPER_USER_PASSWORD || '' }
}

export function isValidIsraeliId(value: string) {
  if (!/^\d{9}$/.test(value)) return false
  return value.split('').reduce((sum, digit, index) => {
    const product = Number(digit) * (index % 2 + 1)
    return sum + (product > 9 ? product - 9 : product)
  }, 0) % 10 === 0
}

function sessionSecret() {
  const codes = accessCodes()
  return process.env.AUTH_SECRET || `${codes.manager}:${codes.superUser}:calendar-access`
}

export function createSessionToken(session: Session) {
  const payload = Buffer.from(JSON.stringify(session)).toString('base64url')
  const signature = createHmac('sha256', sessionSecret()).update(payload).digest('base64url')
  return `${payload}.${signature}`
}

export function readSession(value: string | undefined): Session | null {
  if (!value) return null
  const [payload, signature] = value.split('.')
  if (!payload || !signature) return null
  const expected = createHmac('sha256', sessionSecret()).update(payload).digest('base64url')
  const actual = Buffer.from(signature)
  const target = Buffer.from(expected)
  if (actual.length !== target.length || !timingSafeEqual(actual, target)) return null
  try {
    const session = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8')) as Session
    return session.role && session.userId ? session : null
  } catch {
    return null
  }
}

export function authenticateCode(code: string) {
  const codes = accessCodes()
  const role: UserRole = code === codes.manager ? 'manager' : code === codes.superUser ? 'super_user' : 'user'
  return role === 'user' && !isValidIsraeliId(code) ? null : { role, userId: code }
}
