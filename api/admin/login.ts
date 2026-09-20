import { createHmac, timingSafeEqual } from 'node:crypto'

type VercelRequest = {
  method?: string
  body?: unknown
}

type VercelResponse = {
  status: (code: number) => VercelResponse
  json: (body: unknown) => void
}

function adminToken() {
  const password = process.env.ADMIN_PASSWORD
  if (!password) throw new Error('ADMIN_PASSWORD is missing')
  return createHmac('sha256', password).update('calendar-admin').digest('hex')
}

export function isValidAdminToken(value: string | undefined) {
  if (!value) return false
  const expected = adminToken()
  const actual = Buffer.from(value)
  const target = Buffer.from(expected)
  return actual.length === target.length && timingSafeEqual(actual, target)
}

export default function handler(request: VercelRequest, response: VercelResponse) {
  if (request.method !== 'POST') {
    response.status(405).json({ error: 'Method not allowed' })
    return
  }
  const password = (request.body as { password?: unknown } | undefined)?.password
  if (typeof password !== 'string' || password !== process.env.ADMIN_PASSWORD) {
    response.status(200).json({ valid: false })
    return
  }
  response.status(200).json({ valid: true, token: adminToken() })
}
