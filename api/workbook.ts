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
    response.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.setHeader('Content-Disposition', 'attachment; filename="calendar-events.xlsx"')
    response.send(Buffer.from(await result.arrayBuffer()))
  } catch (error) {
    response.status(500).json({ error: error instanceof Error ? error.message : 'Unexpected server error' })
  }
}
