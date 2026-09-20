import { isValidAdminToken } from './admin/login'

type EventRecord = {
  date: string
  title: string
  className: string
  deleted: boolean
}

type VercelRequest = {
  method?: string
  body?: unknown
  headers?: Record<string, string | undefined>
}

type VercelResponse = {
  status: (code: number) => VercelResponse
  json: (body: unknown) => void
}

function configuration() {
  const url = process.env.SUPABASE_URL
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY
  if (!url || !key) throw new Error('Supabase environment variables are missing')
  return { url, key }
}

async function supabaseRequest(path: string, init: RequestInit = {}) {
  const { url, key } = configuration()
  return fetch(`${url}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: key,
      Authorization: `Bearer ${key}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  })
}

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    if (request.method === 'GET') {
      const result = await supabaseRequest('calendar_events?select=date,title,class_name,deleted')
      if (!result.ok) throw new Error(await result.text())
      const rows = await result.json() as Array<{ date: string; title: string; class_name: string; deleted: boolean }>
      response.status(200).json(rows.map((row) => ({ date: row.date, title: row.title, className: row.class_name, deleted: row.deleted })))
      return
    }

    const payload = request.body as Partial<EventRecord> | undefined
    if (!payload?.date) {
      response.status(400).json({ error: 'Missing event date' })
      return
    }

    if (request.method === 'POST') {
      if (!payload.title) {
        response.status(400).json({ error: 'Missing event title' })
        return
      }
      const result = await supabaseRequest('calendar_events', {
        method: 'POST',
        headers: { Prefer: 'resolution=merge-duplicates' },
        body: JSON.stringify({ date: payload.date, title: payload.title, class_name: payload.className || '', deleted: false }),
      })
      if (!result.ok) throw new Error(await result.text())
      response.status(200).json({ ok: true })
      return
    }

    if (request.method === 'DELETE') {
      const authorization = request.headers?.authorization
      if (!isValidAdminToken(authorization?.replace(/^Bearer\s+/i, ''))) {
        response.status(401).json({ error: 'Admin authorization required' })
        return
      }
      const result = await supabaseRequest(`calendar_events?date=eq.${encodeURIComponent(payload.date)}`, {
        method: 'PATCH',
        body: JSON.stringify({ deleted: true }),
      })
      if (!result.ok) throw new Error(await result.text())
      response.status(200).json({ ok: true })
      return
    }

    response.status(405).json({ error: 'Method not allowed' })
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected server error' })
  }
}
