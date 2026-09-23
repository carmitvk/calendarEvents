import * as XLSX from 'xlsx'

type VercelRequest = { method?: string }
type VercelResponse = { status: (code: number) => VercelResponse; setHeader: (name: string, value: string) => void; send: (body: unknown) => void; json: (body: unknown) => void }

const spreadsheetId = '1V3wMw6tGR1XgiHqOov-nPGG-6CXL-e8o3qvA4uxTX5M'
const sheetGid = '2108382964'

export default async function handler(request: VercelRequest, response: VercelResponse) {
  try {
    if (request.method !== 'GET') {
      response.status(405).json({ error: 'Method not allowed' })
      return
    }
    const result = await fetch(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=xlsx&gid=${sheetGid}&t=${Date.now()}`, { cache: 'no-store' })
    if (!result.ok) throw new Error(`Google Sheet download failed: ${result.status}`)
    const workbook = XLSX.read(await result.arrayBuffer(), { type: 'array', bookVBA: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    if (sheet) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
      const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
      if (headerIndex >= 0) {
        const headers = rows[headerIndex].map((cell) => String(cell || ''))
        const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
        const rowsWithData = rows.slice(headerIndex + 1).filter((row) => titleIndex >= 0
          ? String(row[titleIndex] || '').trim()
          : row.some((cell) => String(cell || '').trim()))
        workbook.Sheets[workbook.SheetNames[0]] = XLSX.utils.aoa_to_sheet([...rows.slice(0, headerIndex + 1), ...rowsWithData])
      }
    }
    const filteredWorkbook = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer', bookVBA: true })
    response.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.setHeader('Content-Disposition', 'attachment; filename="calendar-events.xlsx"')
    response.send(filteredWorkbook)
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected server error' })
  }
}
