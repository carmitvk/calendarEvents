import * as XLSX from 'xlsx'

type VercelRequest = { method?: string }
type VercelResponse = { status: (code: number) => VercelResponse; setHeader: (name: string, value: string) => void; send: (body: unknown) => void; json: (body: unknown) => void }
type EventRow = { date: string; title: string; class_name: string }

function configuration() {
  const url = process.env.SUPABASE_URL
  const key = process.env.SUPABASE_SERVICE_ROLE_KEY
  if (!url || !key) throw new Error('Supabase environment variables are missing')
  return { url, key }
}

async function loadEvents() {
  const { url, key } = configuration()
  const result = await fetch(`${url}/rest/v1/calendar_events?select=date,title,class_name&order=date.asc`, {
    headers: { apikey: key, Authorization: `Bearer ${key}` },
  })
  if (!result.ok) throw new Error(await result.text())
  return await result.json() as EventRow[]
}

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    if (request.method !== 'GET') {
      response.status(405).json({ error: 'Method not allowed' })
      return
    }
    const rows = await loadEvents()
    const sheet = XLSX.utils.json_to_sheet(rows.map((row) => ({
      'תאריך לועזי': row.date,
      'שם החוגגת': row.title,
      'כיתה': row.class_name,
    })))
    sheet['!cols'] = [{ wch: 14 }, { wch: 28 }, { wch: 14 }]
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, sheet, 'אירועים')
    const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', cellDates: true })
    response.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.setHeader('Content-Disposition', 'attachment; filename="calendar-events.xlsx"')
    response.send(output)
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected server error' })
  }
}
