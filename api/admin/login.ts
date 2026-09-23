import { authenticateCode, createSessionToken, type UserRole } from '../auth.js'

type VercelRequest = {
  method?: string
  body?: unknown
}

type VercelResponse = {
  status: (code: number) => VercelResponse
  json: (body: unknown) => void
}

export default async function handler(request: VercelRequest, response: VercelResponse) {
  if (request.method !== 'POST') {
    response.status(405).json({ error: 'Method not allowed' })
    return
  }
  const password = (request.body as { password?: unknown } | undefined)?.password
  if (typeof password !== 'string' || !password) {
    response.status(200).json({ valid: false })
    return
  }
  const session = authenticateCode(password)
  if (!session) {
    response.status(200).json({ valid: false })
    return
  }
  response.status(200).json({ valid: true, ...session, token: await createSessionToken(session) })
}
