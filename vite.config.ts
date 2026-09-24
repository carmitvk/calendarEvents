import fs from 'node:fs'
import path from 'node:path'
import { createHmac } from 'node:crypto'
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'

const primaryWorkbookDir = path.resolve(process.cwd(), 'xmlfiles')
const legacyWorkbookDir = path.resolve(process.cwd(), 'xlsfiles')
const workbookFileName = 'תאריכיבתמצוות.xlsm'
const templateFileName = 'תאריכיבתמצוות-תשפז.xlsm'

fs.mkdirSync(primaryWorkbookDir, { recursive: true })

const workbookPath = path.join(
  fs.existsSync(path.join(primaryWorkbookDir, workbookFileName)) ? primaryWorkbookDir : legacyWorkbookDir,
  workbookFileName,
)
const requestedTemplatePath = path.join(primaryWorkbookDir, templateFileName)
const legacyTemplatePath = path.join(legacyWorkbookDir, templateFileName)
const adminPasswordPath = path.join(primaryWorkbookDir, 'admin-pwd.txt')
const templatePath = fs.existsSync(requestedTemplatePath)
  ? requestedTemplatePath
  : fs.existsSync(legacyTemplatePath)
    ? legacyTemplatePath
    : workbookPath

function isValidIsraeliId(value: string) {
  if (!/^\d{9}$/.test(value)) return false
  return value.split('').reduce((sum, digit, index) => {
    const product = Number(digit) * (index % 2 + 1)
    return sum + (product > 9 ? product - 9 : product)
  }, 0) % 10 === 0
}

function localLogin(password: string) {
  const codes = fs.readFileSync(adminPasswordPath, 'utf8').split(/\r?\n/).map((value) => value.trim()).filter(Boolean)
  const managerCode = codes[0] || ''
  const superUserCode = codes[1] || ''
  const role = password === managerCode ? 'manager' : password === superUserCode ? 'super_user' : 'user'
  if (role === 'user' && !isValidIsraeliId(password)) return null
  const session = { role, userId: password }
  const payload = Buffer.from(JSON.stringify(session)).toString('base64url')
  const secret = process.env.AUTH_SECRET || `${managerCode}:${superUserCode}:calendar-access`
  const signature = createHmac('sha256', secret).update(payload).digest('base64url')
  return { valid: true, ...session, token: `${payload}.${signature}` }
}

if (!fs.existsSync(workbookPath) && fs.existsSync(path.join(legacyWorkbookDir, workbookFileName))) {
  fs.copyFileSync(path.join(legacyWorkbookDir, workbookFileName), workbookPath)
}

function excelSerialFromDateKey(dateKey: string) {
  const [year, month, day] = dateKey.split('-').map(Number)
  return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000
}

function excelDateFromDateKey(dateKey: string) {
  const [year, month, day] = dateKey.split('-').map(Number)
  return new Date(Date.UTC(year, month - 1, day))
}

function excelSerialToDateKey(value: number): string {
  const base = new Date(Date.UTC(1899, 11, 30))
  const date = new Date(base.getTime() + Math.floor(value) * 86400000)
  const year = date.getUTCFullYear()
  const month = String(date.getUTCMonth() + 1).padStart(2, '0')
  const day = String(date.getUTCDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function normalizeDateKey(value: unknown): string {
  if (value instanceof Date) {
    const year = value.getFullYear()
    const month = String(value.getMonth() + 1).padStart(2, '0')
    const day = String(value.getDate()).padStart(2, '0')
    return `${year}-${month}-${day}`
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return ''
    return excelSerialToDateKey(value)
  }

  const text = String(value ?? '').trim()
  if (!text) return ''

  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)
  if (iso) return `${iso[1]}-${String(Number(iso[2])).padStart(2, '0')}-${String(Number(iso[3])).padStart(2, '0')}`

  const eu = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  if (eu) return `${eu[3]}-${String(Number(eu[2])).padStart(2, '0')}-${String(Number(eu[1])).padStart(2, '0')}`

  const us = text.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/)
  if (us) return `${us[1]}-${String(Number(us[2])).padStart(2, '0')}-${String(Number(us[3])).padStart(2, '0')}`

  return ''
}

function escapeXml(value: string) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

function readWorkbookEvents(buffer: Buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', bookVBA: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
  if (!sheet || headerIndex < 0) return []
  const headers = rows[headerIndex].map((cell) => String(cell || ''))
  const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
  const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
  const classIndex = headers.findIndex((header) => header.includes('כיתה'))
  return rows.slice(headerIndex + 1).map((row, index) => ({
    rowNumber: headerIndex + index + 2,
    date: dateIndex >= 0 ? String(row[dateIndex] || '') : '',
    title: titleIndex >= 0 ? String(row[titleIndex] || '') : '',
    className: classIndex >= 0 ? String(row[classIndex] || '') : '',
  })).filter((item) => item.date && item.title)
}

async function updateWorkbookEvents(events: Array<{ date: string; title: string; className: string }>) {
  const workbookBuffer = fs.readFileSync(workbookPath)
  const workbook = XLSX.read(workbookBuffer, { type: 'buffer', bookVBA: true })
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]

  if (!sheet) {
    throw new Error('Workbook sheet not found')
  }

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))

  if (headerIndex < 0) {
    throw new Error('Workbook header not found')
  }

  const headers = rows[headerIndex].map((cell) => String(cell || ''))
  const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
  const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
  const classIndex = headers.findIndex((header) => header.includes('כיתה'))

  if (dateIndex < 0 || titleIndex < 0 || classIndex < 0) {
    throw new Error('Required columns not found in workbook')
  }

  const rawRows = rows.slice(headerIndex + 1)
  const existingDates = new Map<string, number>()
  rawRows.forEach((row, index) => {
    const dateKey = normalizeDateKey(row[dateIndex])
    if (dateKey) {
      existingDates.set(dateKey, headerIndex + 1 + index)
      const dateCell = XLSX.utils.encode_cell({ r: headerIndex + 1 + index, c: dateIndex })
      sheet[dateCell] = { ...(sheet[dateCell] || {}), t: 'd', v: excelDateFromDateKey(dateKey), z: 'dd/mm/yyyy' }
    }
  })

  let nextRow = headerIndex + 1 + rawRows.length
  events.forEach((event) => {
    const eventDateKey = normalizeDateKey(event.date)
    const rowIndex = existingDates.get(eventDateKey)
    const targetRow = typeof rowIndex === 'number' ? rowIndex : nextRow

    if (typeof rowIndex !== 'number') {
      nextRow += 1
      existingDates.set(eventDateKey, targetRow)
    }

    const dateCell = XLSX.utils.encode_cell({ r: targetRow, c: dateIndex })
    sheet[dateCell] = { t: 'd', v: excelDateFromDateKey(eventDateKey), z: 'dd/mm/yyyy' }

    const titleCell = XLSX.utils.encode_cell({ r: targetRow, c: titleIndex })
    sheet[titleCell] = { t: 's', v: event.title }

    const classCell = XLSX.utils.encode_cell({ r: targetRow, c: classIndex })
    sheet[classCell] = { t: 's', v: event.className || '' }
  })

  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { c: 0, r: 0 }, e: { c: classIndex, r: nextRow } }
  range.e.r = Math.max(range.e.r, nextRow)
  sheet['!ref'] = XLSX.utils.encode_range(range)
  sheet['!cols'] = sheet['!cols'] || []
  sheet['!cols'][dateIndex] = { ...(sheet['!cols'][dateIndex] || {}), wch: 14 }
  sheet['!cols'][titleIndex] = { ...(sheet['!cols'][titleIndex] || {}), wch: 24 }
  sheet['!cols'][classIndex] = { ...(sheet['!cols'][classIndex] || {}), wch: 12 }

  const out = XLSX.write(workbook, {
    bookType: 'xlsm',
    type: 'buffer',
    bookVBA: true,
    compression: true,
    cellDates: true,
  })
  fs.writeFileSync(workbookPath, out)
}

async function deleteWorkbookEvent(date: string) {
  const workbook = XLSX.read(fs.readFileSync(workbookPath), { type: 'buffer', bookVBA: true })
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  if (!sheet) throw new Error('Workbook sheet not found')

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
  if (headerIndex < 0) throw new Error('Workbook header not found')

  const headers = rows[headerIndex].map((cell) => String(cell || ''))
  const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
  const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
  const classIndex = headers.findIndex((header) => header.includes('כיתה'))
  const targetIndex = rows.slice(headerIndex + 1).findIndex((row) => normalizeDateKey(row[dateIndex]) === normalizeDateKey(date))
  if (dateIndex < 0 || titleIndex < 0 || classIndex < 0 || targetIndex < 0) return

  rows.slice(headerIndex + 1).forEach((row, index) => {
    const dateKey = normalizeDateKey(row[dateIndex])
    if (!dateKey) return
    const dateCell = XLSX.utils.encode_cell({ r: headerIndex + 1 + index, c: dateIndex })
    sheet[dateCell] = { ...(sheet[dateCell] || {}), t: 'd', v: excelDateFromDateKey(dateKey), z: 'dd/mm/yyyy' }
  })

  const rowIndex = headerIndex + 1 + targetIndex
  sheet[XLSX.utils.encode_cell({ r: rowIndex, c: titleIndex })] = { t: 's', v: '' }
  sheet[XLSX.utils.encode_cell({ r: rowIndex, c: classIndex })] = { t: 's', v: '' }

  const out = XLSX.write(workbook, {
    bookType: 'xlsm',
    type: 'buffer',
    bookVBA: true,
    compression: true,
    cellDates: true,
  })
  fs.writeFileSync(workbookPath, out)
}

function workbookApi() {
  return {
    name: 'workbook-api',
    configureServer(server: { middlewares: { use: (handler: (request: any, response: any, next: () => void) => void) => void } }) {
      void updateWorkbookEvents([]).catch(() => undefined)
      server.middlewares.use((request, response, next) => {
        if (request.url === '/api/admin/login' && request.method === 'POST') {
          let body = ''
          request.on('data', (chunk: Buffer) => { body += chunk.toString() })
          request.on('error', (error: Error) => {
            response.statusCode = 400
            response.end(JSON.stringify({ error: error.message }))
          })
          request.on('end', () => {
            try {
              const payload = JSON.parse(body) as { password?: string }
              const result = typeof payload.password === 'string' ? localLogin(payload.password) : null
              response.setHeader('Content-Type', 'application/json')
              response.end(JSON.stringify(result || { valid: false }))
            } catch {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Invalid password request' }))
            }
          })
          return
        }
        if (request.url === '/api/workbook' && request.method === 'GET') {
          response.setHeader('Content-Type', 'application/vnd.ms-excel.sheet.macroEnabled.12')
          response.end(fs.readFileSync(workbookPath))
          return
        }
        if (request.url === '/api/events' && request.method === 'GET') {
          try {
            response.setHeader('Content-Type', 'application/json')
            response.end(JSON.stringify(readWorkbookEvents(fs.readFileSync(workbookPath)).map((event, index) => ({ ...event, id: index + 1 }))))
          } catch (error) {
            response.statusCode = 500
            response.end(JSON.stringify({ error: error instanceof Error ? error.message : 'Workbook read failed' }))
          }
          return
        }
        if (request.url === '/api/events' && request.method === 'POST') {
          let body = ''
          request.on('data', (chunk: Buffer) => { body += chunk.toString() })
          request.on('end', () => {
            try {
              const payload = JSON.parse(body) as { date?: string; title?: string; className?: string }
              if (!payload.date || !payload.title || !payload.className) {
                response.statusCode = 400
                response.end(JSON.stringify({ error: 'Missing event data' }))
                return
              }
              void updateWorkbookEvents([{ date: payload.date, title: payload.title, className: payload.className }])
                .then(() => {
                  response.statusCode = 201
                  response.setHeader('Content-Type', 'application/json')
                  response.end(JSON.stringify({ ok: true }))
                })
                .catch((error: unknown) => {
                  response.statusCode = 500
                  response.end(JSON.stringify({ error: error instanceof Error ? error.message : 'Workbook update failed' }))
                })
            } catch {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Invalid JSON' }))
            }
          })
          return
        }
        if (request.url === '/api/events' && request.method === 'DELETE') {
          let body = ''
          request.on('data', (chunk: Buffer) => { body += chunk.toString() })
          request.on('end', () => {
            try {
              const payload = JSON.parse(body) as { date?: string }
              if (!payload.date) {
                response.statusCode = 400
                response.end(JSON.stringify({ error: 'Missing event date' }))
                return
              }
              void deleteWorkbookEvent(payload.date)
                .then(() => {
                  response.statusCode = 200
                  response.setHeader('Content-Type', 'application/json')
                  response.end(JSON.stringify({ ok: true }))
                })
                .catch((error: unknown) => {
                  response.statusCode = 500
                  response.end(JSON.stringify({ error: error instanceof Error ? error.message : 'Workbook update failed' }))
                })
            } catch {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Invalid JSON' }))
            }
          })
          return
        }
        if (request.url === '/api/workbook/event' && request.method === 'POST') {
          let body = ''
          request.on('data', (chunk: Buffer) => { body += chunk.toString() })
          request.on('error', (error: Error) => {
            response.statusCode = 400
            response.end(JSON.stringify({ error: error.message }))
          })
          request.on('end', () => {
            let payload: { date?: string; title?: string; className?: string }
            try {
              payload = JSON.parse(body)
            } catch {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Invalid JSON' }))
              return
            }
            if (!payload.date || !payload.title) {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Missing event data' }))
              return
            }
            response.statusCode = 202
            response.end()
            void updateWorkbookEvents([{ date: payload.date, title: payload.title, className: payload.className || '' }])
              .catch((error: unknown) => console.error('Workbook update failed:', error))
          })
          return
        }
        if (request.url === '/api/workbook/event' && request.method === 'DELETE') {
          let body = ''
          request.on('data', (chunk: Buffer) => { body += chunk.toString() })
          request.on('error', (error: Error) => {
            response.statusCode = 400
            response.end(JSON.stringify({ error: error.message }))
          })
          request.on('end', () => {
            let payload: { date?: string }
            try {
              payload = JSON.parse(body)
            } catch {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Invalid JSON' }))
              return
            }
            if (!payload.date) {
              response.statusCode = 400
              response.end(JSON.stringify({ error: 'Missing event date' }))
              return
            }
            void deleteWorkbookEvent(payload.date)
              .then(() => {
                response.statusCode = 204
                response.end()
              })
              .catch((error: unknown) => {
                response.statusCode = 500
                response.end(JSON.stringify({ error: error instanceof Error ? error.message : 'Workbook update failed' }))
              })
          })
          return
        }
        next()
      })
    },
  }
}

export default defineConfig({
  plugins: [workbookApi(), react()],
})