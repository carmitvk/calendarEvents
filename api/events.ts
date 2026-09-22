import * as XLSX from 'xlsx'

type VercelRequest = { method?: string; body?: unknown }
type VercelResponse = { status: (code: number) => VercelResponse; json: (body: unknown) => void }

const spreadsheetId = '1V3wMw6tGR1XgiHqOov-nPGG-6CXL-e8o3qvA4uxTX5M'
const sheetGid = '2108382964'

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
  return rows.slice(headerIndex + 1).map((row, index) => ({
    id: index + 1,
    date: normalizeDate(row[dateIndex]),
    title: String(row[titleIndex] || '').trim(),
    className: String(row[classIndex] || '').trim(),
  })).filter((event) => event.date && event.title)
}

async function forwardWrite(payload: unknown) {
  const scriptUrl = process.env.GOOGLE_SHEETS_SCRIPT_URL
  const token = process.env.GOOGLE_SHEETS_SCRIPT_TOKEN
  if (!scriptUrl || !token) throw new Error('Google Sheets write connection is not configured')
  const response = await fetch(scriptUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ...(payload as object), token }),
  })
  if (!response.ok) throw new Error(`Google Sheets update failed: ${response.status}`)
  const result = await response.json() as { ok?: boolean; error?: string }
  if (!result.ok) throw new Error(result.error || 'Google Sheets update failed')
}

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    if (request.method === 'GET') {
      response.status(200).json(await loadSheetEvents())
      return
    }
    if (request.method === 'POST') {
      await forwardWrite({ action: 'upsert', event: request.body })
      response.status(201).json({ ok: true })
      return
    }
    if (request.method === 'DELETE') {
      const payload = request.body as { date?: string } | undefined
      await forwardWrite({ action: 'delete', date: payload?.date })
      response.status(200).json({ ok: true })
      return
    }
    response.status(405).json({ error: 'Method not allowed' })
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected error' })
  }
}
