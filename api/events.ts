import * as XLSX from 'xlsx'
import { readSession } from './auth.js'

type VercelRequest = { method?: string; body?: unknown; headers?: Record<string, string | undefined> }
type VercelResponse = { status: (code: number) => VercelResponse; json: (body: unknown) => void }

const spreadsheetId = '1V3wMw6tGR1XgiHqOov-nPGG-6CXL-e8o3qvA4uxTX5M'
const sheetGid = '2108382964'
const sharedClasses = new Set(['שכבתי', 'בית ספרי', 'בנות השכבה'])

function normalizeDate(value: unknown) {
  if (value instanceof Date) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value)
    return date ? `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}` : ''
  }
  const text = String(value || '').trim()
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  return match ? `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}` : text.slice(0, 10)
}

async function loadSheetEvents() {
  const response = await fetch(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${sheetGid}&t=${Date.now()}`, { cache: 'no-store' })
  if (!response.ok) throw new Error(`Google Sheet download failed: ${response.status}`)
  const workbook = XLSX.read(await response.arrayBuffer(), { type: 'array' })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  if (!sheet) throw new Error('Google Sheet tab not found')
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
  if (headerIndex < 0) throw new Error('Google Sheet header not found')
  const headers = rows[headerIndex].map((cell) => String(cell || ''))
  const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
  const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
  const classIndex = headers.findIndex((header) => header.includes('כיתה'))
  const ownerIndex = headers.findIndex((header) => header.includes('תז') || header.includes('ת.ז') || header.includes('ת״ז'))
  return rows.slice(headerIndex + 1).map((row, index) => ({
    id: index + 1,
    date: normalizeDate(row[dateIndex]),
    title: String(row[titleIndex] || '').trim(),
    className: String(row[classIndex] || '').trim(),
    ownerId: ownerIndex >= 0 ? String(row[ownerIndex] || '').trim() : '',
  })).filter((event) => event.date && event.title)
}

async function requestSession(request: VercelRequest) {
  const authorization = request.headers?.authorization || request.headers?.Authorization
  return readSession(authorization?.replace(/^Bearer\s+/i, ''))
}

async function forwardWrite(payload: unknown) {
  const scriptUrl = process.env.GOOGLE_SHEETS_SCRIPT_URL || process.env.TS_SCRIPT_URL
  const token = process.env.GOOGLE_SHEETS_SCRIPT_TOKEN || process.env.WRITE_TOKEN
  if (!scriptUrl || !token) throw new Error('Google Sheets write connection is not configured')
  const response = await fetch(scriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...(payload as object), token }),
  })
  const responseText = await response.text()
  let result: { ok?: boolean; error?: string }
  try {
    result = JSON.parse(responseText) as { ok?: boolean; error?: string }
  } catch {
    const detail = responseText.replace(/\s+/g, ' ').trim().slice(0, 240)
    throw new Error(`Google Apps Script returned an invalid response (${response.status}): ${detail || 'empty response'}`)
  }
  if (!response.ok) throw new Error(`Google Sheets update failed: ${response.status} ${result.error || ''}`.trim())
  if (!result.ok) throw new Error(result.error || 'Google Sheets update failed')
}

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    if (request.method === 'GET') {
      response.status(200).json(await loadSheetEvents())
      return
    }
    if (request.method === 'POST') {
      const session = await requestSession(request)
      if (!session) {
        response.status(401).json({ error: 'Login required' })
        return
      }
      const event = request.body as { title?: string; className?: string; date?: string } | undefined
      if (!event?.title || !event.className || !event.date) {
        response.status(400).json({ error: 'Incomplete event' })
        return
      }
      if (sharedClasses.has(event.className) && session.role === 'user') {
        response.status(403).json({ error: 'Only managers can create shared events' })
        return
      }
      const existingEvents = await loadSheetEvents()
      const existingUserEvent = session.role === 'user' ? existingEvents.find((item) => item.ownerId === session.userId) : undefined
      if (existingUserEvent) {
        response.status(403).json({ error: 'A regular user may create only one event', existingDate: existingUserEvent.date })
        return
      }
      await forwardWrite({ action: 'upsert', event: { ...event, ownerId: session.userId }, role: session.role, userId: session.userId })
      response.status(201).json({ ok: true })
      return
    }
    if (request.method === 'DELETE') {
      const session = await requestSession(request)
      if (!session) {
        response.status(401).json({ error: 'Login required' })
        return
      }
      const payload = request.body as { date?: string } | undefined
      const existingEvent = (await loadSheetEvents()).find((item) => item.date === payload?.date)
      if (!existingEvent) {
        response.status(404).json({ error: 'Event not found' })
        return
      }
      if (session.role !== 'super_user' && existingEvent.ownerId !== session.userId) {
        response.status(403).json({ error: 'You can delete only your own event' })
        return
      }
      await forwardWrite({ action: 'delete', date: payload?.date, role: session.role, userId: session.userId })
      response.status(200).json({ ok: true })
      return
    }
    response.status(405).json({ error: 'Method not allowed' })
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected error' })
  }
}
