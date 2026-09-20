import * as XLSX from 'xlsx'
import { isValidAdminToken } from './admin/login'

type VercelRequest = {
  method?: string
  body?: unknown
  headers?: Record<string, string | undefined>
}

type VercelResponse = {
  status: (code: number) => VercelResponse
  setHeader: (name: string, value: string) => void
  send: (body: unknown) => void
  json: (body: unknown) => void
}

function configuration() {
  const url = process.env.SUPABASE_URL
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY
  const bucket = process.env.SUPABASE_WORKBOOK_BUCKET || 'workbooks'
  const filePath = process.env.SUPABASE_WORKBOOK_PATH || 'calendar-events.xlsm'
  if (!url || !key) throw new Error('Supabase environment variables are missing')
  return { url, key, bucket, filePath }
}

function storageUrl() {
  const { url, bucket, filePath } = configuration()
  return `${url}/storage/v1/object/${encodeURIComponent(bucket)}/${filePath.split('/').map(encodeURIComponent).join('/')}`
}

function storageHeaders() {
  const { key } = configuration()
  return { apikey: key, Authorization: `Bearer ${key}` }
}

async function downloadWorkbook() {
  const result = await fetch(storageUrl(), { headers: storageHeaders() })
  if (!result.ok) throw new Error(`Workbook download failed: ${result.status}`)
  return Buffer.from(await result.arrayBuffer())
}

async function uploadWorkbook(workbook: XLSX.WorkBook) {
  const { filePath } = configuration()
  const data = XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer', bookVBA: true, compression: true, cellDates: true })
  const result = await fetch(storageUrl(), {
    method: 'POST',
    headers: { ...storageHeaders(), 'Content-Type': 'application/vnd.ms-excel.sheet.macroEnabled.12', 'x-upsert': 'true' },
    body: data,
  })
  if (!result.ok) throw new Error(`Workbook upload failed for ${filePath}: ${await result.text()}`)
}

function normalizeDate(value: unknown) {
  if (value instanceof Date) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`
  }
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value)
    return date ? `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}` : ''
  }
  const text = String(value ?? '').trim()
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  return match ? `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}` : text.slice(0, 10)
}

function columns(sheet: XLSX.WorkSheet) {
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
  if (headerIndex < 0) throw new Error('Workbook header not found')
  const headers = rows[headerIndex].map((cell) => String(cell || ''))
  const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
  const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
  const classIndex = headers.findIndex((header) => header.includes('כיתה'))
  if (dateIndex < 0 || titleIndex < 0 || classIndex < 0) throw new Error('Required workbook columns not found')
  return { rows, headerIndex, dateIndex, titleIndex, classIndex }
}

function excelDate(date: string) {
  const [year, month, day] = date.split('-').map(Number)
  return new Date(Date.UTC(year, month - 1, day))
}

async function updateEvent(payload: { date: string; title: string; className?: string }) {
  const workbook = XLSX.read(await downloadWorkbook(), { type: 'buffer', bookVBA: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  if (!sheet) throw new Error('Workbook sheet not found')
  const { rows, headerIndex, dateIndex, titleIndex, classIndex } = columns(sheet)
  const existingRows = rows.slice(headerIndex + 1)
  const targetOffset = existingRows.findIndex((row) => normalizeDate(row[dateIndex]) === payload.date)
  const targetRow = targetOffset >= 0 ? headerIndex + 1 + targetOffset : headerIndex + 1 + existingRows.length
  sheet[XLSX.utils.encode_cell({ r: targetRow, c: dateIndex })] = { t: 'd', v: excelDate(payload.date), z: 'dd/mm/yyyy' }
  sheet[XLSX.utils.encode_cell({ r: targetRow, c: titleIndex })] = { t: 's', v: payload.title }
  sheet[XLSX.utils.encode_cell({ r: targetRow, c: classIndex })] = { t: 's', v: payload.className || '' }
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { c: 0, r: 0 }, e: { c: classIndex, r: targetRow } }
  range.e.r = Math.max(range.e.r, targetRow)
  sheet['!ref'] = XLSX.utils.encode_range(range)
  await uploadWorkbook(workbook)
}

async function deleteEvent(date: string) {
  const workbook = XLSX.read(await downloadWorkbook(), { type: 'buffer', bookVBA: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  if (!sheet) throw new Error('Workbook sheet not found')
  const { rows, headerIndex, dateIndex, titleIndex, classIndex } = columns(sheet)
  const targetOffset = rows.slice(headerIndex + 1).findIndex((row) => normalizeDate(row[dateIndex]) === date)
  if (targetOffset < 0) return
  const row = headerIndex + 1 + targetOffset
  sheet[XLSX.utils.encode_cell({ r: row, c: titleIndex })] = { t: 's', v: '' }
  sheet[XLSX.utils.encode_cell({ r: row, c: classIndex })] = { t: 's', v: '' }
  await uploadWorkbook(workbook)
}

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    if (request.method === 'GET') {
      response.setHeader('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12')
      response.send(await downloadWorkbook())
      return
    }

    const payload = request.body as { date?: string; title?: string; className?: string } | undefined
    if (!payload?.date) {
      response.status(400).json({ error: 'Missing event date' })
      return
    }
    if (request.method === 'POST') {
      if (!payload.title) {
        response.status(400).json({ error: 'Missing event title' })
        return
      }
      await updateEvent({ date: payload.date, title: payload.title, className: payload.className })
      response.status(200).json({ ok: true })
      return
    }
    if (request.method === 'DELETE') {
      const authorization = request.headers?.authorization
      if (!isValidAdminToken(authorization?.replace(/^Bearer\s+/i, ''))) {
        response.status(401).json({ error: 'Admin authorization required' })
        return
      }
      await deleteEvent(payload.date)
      response.status(200).json({ ok: true })
      return
    }
    response.status(405).json({ error: 'Method not allowed' })
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected server error' })
  }
}
