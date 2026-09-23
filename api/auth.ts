export type UserRole = 'user' | 'manager' | 'super_user'
export type Session = { role: UserRole; userId: string }

function accessCodes() {
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

function encode(value: string) {
  const bytes = new TextEncoder().encode(value)
  let binary = ''
  bytes.forEach((byte) => { binary += String.fromCharCode(byte) })
  return btoa(binary).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

function decode(value: string) {
  const padded = value.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((value.length + 3) % 4)
  const binary = atob(padded)
  return new TextDecoder().decode(Uint8Array.from(binary, (character) => character.charCodeAt(0)))
}

async function signature(value: string) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(sessionSecret()), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign'])
  const bytes = new Uint8Array(await crypto.subtle.sign('HMAC', key, new TextEncoder().encode(value)))
  return encode(String.fromCharCode(...bytes))
}

export async function createSessionToken(session: Session) {
  const payload = encode(JSON.stringify(session))
  return `${payload}.${await signature(payload)}`
}

export async function readSession(value: string | undefined): Promise<Session | null> {
  if (!value) return null
  const [payload, suppliedSignature] = value.split('.')
  if (!payload || !suppliedSignature || suppliedSignature !== await signature(payload)) return null
  try {
    const session = JSON.parse(decode(payload)) as Session
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
