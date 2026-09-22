import { ChangeEvent, FormEvent, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

type EventItem = {
  id: number
  title: string
  className: string
  date: string
}

type WorkbookState = {
  workbook: XLSX.WorkBook
  fileName: string
}

type ExcelFileHandle = {
  getFile: () => Promise<File>
  createWritable: () => Promise<{ write: (data: ArrayBuffer) => Promise<void>; close: () => Promise<void> }>
}

const monthNames = [
  'ינואר', 'פברואר', 'מרץ', 'אפריל', 'מאי', 'יוני',
  'יולי', 'אוגוסט', 'ספטמבר', 'אוקטובר', 'נובמבר', 'דצמבר',
]
const dayNames = ['א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ש']

const today = new Date()
today.setHours(0, 0, 0, 0)

function toDateKey(date: Date) {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function formatDateForDisplay(dateKey: string) {
  const [year, month, day] = dateKey.split('-')
  return year && month && day ? `${day}/${month}/${year}` : dateKey
}

function toHebrewNumeral(number: number) {
  const units = ['', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט']
  const tens = ['', 'י', 'כ', 'ל']
  if (number === 15) return 'טו'
  if (number === 16) return 'טז'
  if (number < 10) return units[number]
  if (number < 40) return `${tens[Math.floor(number / 10)]}${units[number % 10]}`
  return `${'מ'.repeat(Math.floor(number / 40))}${units[number % 10]}`
}

function getHebrewDate(date: Date) {
  const stableDate = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), 12))
  const parts = new Intl.DateTimeFormat('he-IL-u-ca-hebrew', {
    day: 'numeric',
    month: 'long',
    numberingSystem: 'latn',
    timeZone: 'UTC',
  }).formatToParts(stableDate)
  const dayValue = Number(parts.find((part) => part.type === 'day')?.value || 0)
  const day = toHebrewNumeral(dayValue)
  const month = parts.find((part) => part.type === 'month')?.value || ''
  return `${day} ${month}`
}

function toHebrewYear(year: number) {
  let remainder = year % 1000
  const values: Array<[number, string]> = [[400, 'ת'], [300, 'ש'], [200, 'ר'], [100, 'ק']]
  let result = ''
  values.forEach(([value, letter]) => {
    while (remainder >= value) {
      result += letter
      remainder -= value
    }
  })
  if (remainder === 15) {
    result += 'טו'
  } else if (remainder === 16) {
    result += 'טז'
  } else {
    const tens = ['', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ']
    const units = ['', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט']
    result += tens[Math.floor(remainder / 10)] || ''
    result += units[remainder % 10] || ''
  }
  if (result.length > 1) return `${result.slice(0, -1)}״${result.slice(-1)}`
  return `${result}׳`
}

function getHebrewYearRange(date: Date) {
  const startDate = new Date(date.getFullYear(), date.getMonth(), 1)
  const endDate = new Date(date.getFullYear(), date.getMonth() + 1, 0)
  const formatter = new Intl.DateTimeFormat('he-IL-u-ca-hebrew', {
    year: 'numeric',
    numberingSystem: 'latn',
    timeZone: 'UTC',
  })
  const startYear = Number(formatter.format(startDate).match(/\d+/)?.[0] || 0)
  const endYear = Number(formatter.format(endDate).match(/\d+/)?.[0] || startYear)
  const startLabel = toHebrewYear(startYear)
  const endLabel = toHebrewYear(endYear)
  return startYear === endYear ? startLabel : `${startLabel}-${endLabel}`
}

function normalizeExcelDate(value: unknown) {
  if (value instanceof Date) return toDateKey(value)
  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value)
    return date ? `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}` : ''
  }
  const text = String(value || '').trim()
  if (!text) return ''
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/)
  if (match) {
    const first = Number(match[1])
    const second = Number(match[2])
    if (first > 12 && second <= 12) return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`
    if (second > 12 && first <= 12) return `${match[3]}-${match[1].padStart(2, '0')}-${match[2].padStart(2, '0')}`
    return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`
  }
  return text.slice(0, 10)
}

function excelSerialFromDateKey(dateKey: string) {
  const [year, month, day] = dateKey.split('-').map(Number)
  return (Date.UTC(year, month - 1, day) - Date.UTC(1899, 11, 30)) / 86400000
}

function excelDateFromDateKey(dateKey: string) {
  const [year, month, day] = dateKey.split('-').map(Number)
  return new Date(Date.UTC(year, month - 1, day))
}

function findColumn(row: Record<string, unknown>, names: string[]) {
  const key = Object.keys(row).find((candidate) => names.some((name) => candidate.includes(name)))
  return key ? row[key] : ''
}

function readEventsFromWorkbook(workbook: XLSX.WorkBook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]]
  if (!sheet) return []
  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
  const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
  if (headerIndex < 0) return []
  const headers = rows[headerIndex].map((cell) => String(cell || ''))
  const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
  const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
  const classIndex = headers.findIndex((header) => header.includes('כיתה'))
  return rows.slice(headerIndex + 1).map((row, index) => ({
    id: Date.now() + index,
    title: titleIndex >= 0 ? String(row[titleIndex] || '').trim() : '',
    className: classIndex >= 0 ? String(row[classIndex] || '').trim() : '',
    date: dateIndex >= 0 ? normalizeExcelDate(row[dateIndex]) : '',
  })).filter((item) => item.title && item.date)
}

function EventDatePicker({
  selectedDate,
  onSelect,
  events,
  showError,
}: {
  selectedDate: string
  onSelect: (date: string) => void
  events: EventItem[]
  showError: boolean
}) {
  const [pickerMonth, setPickerMonth] = useState(new Date(today.getFullYear(), today.getMonth(), 1))
  const [isOpen, setIsOpen] = useState(false)
  const dateInputRef = useRef<HTMLInputElement>(null)
  const firstDay = new Date(pickerMonth.getFullYear(), pickerMonth.getMonth(), 1)
  const dayOffset = firstDay.getDay()
  const numberOfDays = new Date(pickerMonth.getFullYear(), pickerMonth.getMonth() + 1, 0).getDate()
  const days: (Date | null)[] = []

  for (let index = 0; index < dayOffset; index += 1) days.push(null)
  for (let day = 1; day <= numberOfDays; day += 1) {
    days.push(new Date(pickerMonth.getFullYear(), pickerMonth.getMonth(), day))
  }
  while (days.length % 7 !== 0) days.push(null)

  function moveMonth(amount: number) {
    const nextMonth = new Date(pickerMonth.getFullYear(), pickerMonth.getMonth() + amount, 1)
    const currentMonth = new Date(today.getFullYear(), today.getMonth(), 1)
    if (nextMonth >= currentMonth) setPickerMonth(nextMonth)
  }

  function selectDate(date: string) {
    if (events.some((event) => event.date === date)) return
    onSelect(date)
    dateInputRef.current?.setCustomValidity('')
    setIsOpen(false)
  }

  return (
    <div className="event-date-picker">
      <input
        ref={dateInputRef}
        type="hidden"
        name="date"
        value={selectedDate}
        required
      />
      <button type="button" className="date-display" onClick={() => setIsOpen((current) => !current)} aria-expanded={isOpen}>
        <span>{selectedDate ? formatDateForDisplay(selectedDate) : 'בחרי תאריך'}</span>
        <span className="date-display-arrow">⌄</span>
      </button>
      {showError && <span className="field-error">שדה חובה למילוי</span>}
      {isOpen && <div className="picker-calendar" onPointerDown={(event) => event.stopPropagation()} onClick={(event) => event.stopPropagation()}>
        <div className="picker-header">
        <button type="button" className="picker-arrow" onClick={() => moveMonth(1)} aria-label="חודש הבא" title="חודש הבא">‹</button>
        <strong>{monthNames[pickerMonth.getMonth()]} {pickerMonth.getFullYear()}</strong>
        <button type="button" className="picker-arrow" onClick={() => moveMonth(-1)} disabled={pickerMonth.getFullYear() === today.getFullYear() && pickerMonth.getMonth() === today.getMonth()} aria-label="חודש קודם" title="חודש קודם">›</button>
        </div>
      <div className="picker-week-row">{dayNames.map((day) => <span key={day}>{day}</span>)}</div>
      <div className="picker-grid">
        {days.map((date, index) => {
          if (!date) return <span className="picker-day empty" key={`empty-${index}`} />
          const dateKey = toDateKey(date)
          const isPast = date < today
          const isOccupied = events.some((event) => event.date === dateKey)
          const isDisabled = isPast || isOccupied
          return (
            <button
              type="button"
              className={`picker-day ${isOccupied ? 'occupied' : 'available'} ${selectedDate === dateKey ? 'selected' : ''}`}
              key={dateKey}
              disabled={isDisabled}
              onTouchEnd={(event) => {
                event.preventDefault()
                event.stopPropagation()
                selectDate(dateKey)
              }}
              onClick={() => selectDate(dateKey)}
              title={isOccupied ? 'התאריך כבר תפוס' : undefined}
            >
              <span className="picker-gregorian">{date.getDate()}</span>
              <span className="picker-hebrew">{getHebrewDate(date)}</span>
            </button>
          )
        })}
      </div>
      </div>}
    </div>
  )
}

function App() {
  const [visibleMonth, setVisibleMonth] = useState(new Date(today.getFullYear(), today.getMonth(), 1))
  const [isFormOpen, setIsFormOpen] = useState(false)
  const [events, setEvents] = useState<EventItem[]>([])
  const [selectedEvent, setSelectedEvent] = useState<EventItem | null>(null)
  const [selectedEventDate, setSelectedEventDate] = useState('')
  const [workbookState, setWorkbookState] = useState<WorkbookState | null>(null)
  const [excelFileHandle, setExcelFileHandle] = useState<ExcelFileHandle | null>(null)
  const [dateValidationAttempted, setDateValidationAttempted] = useState(false)
  const [validationErrors, setValidationErrors] = useState({ title: false, className: false })
  const [isAdmin, setIsAdmin] = useState(false)
  const [adminToken, setAdminToken] = useState('')
  const [deleteCandidateDate, setDeleteCandidateDate] = useState('')
  const [isPasswordDialogOpen, setIsPasswordDialogOpen] = useState(false)
  const [adminPassword, setAdminPassword] = useState('')
  const [showAdminPassword, setShowAdminPassword] = useState(false)
  const [passwordError, setPasswordError] = useState(false)
  const [deleteConfirmDate, setDeleteConfirmDate] = useState('')
  const fileInputRef = useRef<HTMLInputElement>(null)

  function applyWorkbook(workbook: XLSX.WorkBook, fileName: string) {
    setEvents(readEventsFromWorkbook(workbook))
    setWorkbookState({ workbook, fileName })
  }

  useEffect(() => {
    function loadSharedEvents() {
      return fetch(`/api/events?t=${Date.now()}`, { cache: 'no-store' })
      .then((response) => response.ok ? response.json() as Promise<EventItem[]> : Promise.reject(new Error('Events not found')))
      .then((loadedEvents) => setEvents(loadedEvents.map((event, index) => ({ ...event, id: index + 1 }))))
      .catch(() => undefined)
    }

    void loadSharedEvents()
    const refreshTimer = window.setInterval(() => { void loadSharedEvents() }, 3000)
    return () => window.clearInterval(refreshTimer)
  }, [])

  const calendarDays = useMemo(() => {
    const firstDay = new Date(visibleMonth.getFullYear(), visibleMonth.getMonth(), 1)
    const dayOffset = firstDay.getDay()
    const numberOfDays = new Date(visibleMonth.getFullYear(), visibleMonth.getMonth() + 1, 0).getDate()
    const days = []

    for (let index = 0; index < dayOffset; index += 1) days.push(null)
    for (let day = 1; day <= numberOfDays; day += 1) {
      days.push(new Date(visibleMonth.getFullYear(), visibleMonth.getMonth(), day))
    }
    while (days.length % 7 !== 0) days.push(null)
    return days
  }, [visibleMonth])

  function moveMonth(amount: number) {
    const nextMonth = new Date(visibleMonth.getFullYear(), visibleMonth.getMonth() + amount, 1)
    const currentMonth = new Date(today.getFullYear(), today.getMonth(), 1)
    if (nextMonth >= currentMonth) setVisibleMonth(nextMonth)
  }

  function addEvent(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    const form = new FormData(event.currentTarget)
    const title = String(form.get('title') || '').trim()
    const className = String(form.get('className') || '').trim()
    const date = String(form.get('date') || '')
    setValidationErrors({ title: !title, className: !className })
    setDateValidationAttempted(!date)
    if (!title || !className || !date || events.some((event) => event.date === date)) return

    const newEvent = { id: Date.now(), title, className, date }
    const nextEvents = [...events, newEvent]
    setEvents(nextEvents)
    void fetch('/api/events', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(newEvent),
    }).then((response) => {
      if (!response.ok) throw new Error('Event could not be saved')
    }).catch(() => {
      setEvents(events)
      window.alert('לא ניתן לשמור את האירוע. נסי שוב בעוד רגע.')
    })
    setSelectedEventDate('')
    setIsFormOpen(false)
    event.currentTarget.reset()
  }

  function requestAdminAccess() {
    if (isAdmin) {
      setIsAdmin(false)
      setDeleteCandidateDate('')
      return
    }
    setAdminPassword('')
    setShowAdminPassword(false)
    setPasswordError(false)
    setIsPasswordDialogOpen(true)
  }

  async function verifyAdminPassword(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()
    const response = await fetch('/api/admin/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ password: adminPassword }),
    })
    const result = response.ok ? await response.json() as { valid?: boolean; token?: string } : { valid: false }
    if (result.valid) {
      setIsAdmin(true)
      setAdminToken(result.token || '')
      setIsPasswordDialogOpen(false)
      setAdminPassword('')
      setShowAdminPassword(false)
      setPasswordError(false)
      return
    }
    setPasswordError(true)
  }

  async function deleteEvent(date: string) {
    if (!isAdmin) return
    const previousEvents = events
    setEvents((current) => current.filter((event) => event.date !== date))
    setDeleteCandidateDate('')
    setDeleteConfirmDate('')
    const response = await fetch('/api/events', {
      method: 'DELETE',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${adminToken}` },
      body: JSON.stringify({ date }),
    })
    if (!response.ok) setEvents(previousEvents)
  }

  function loadWorkbook(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = () => {
      const workbook = XLSX.read(reader.result, { type: 'array', bookVBA: true })
      applyWorkbook(workbook, file.name)
      setExcelFileHandle(null)
      event.target.value = ''
    }
    reader.readAsArrayBuffer(file)
  }

  async function openWorkbookDirectly() {
    const picker = (window as Window & {
      showOpenFilePicker?: (options?: unknown) => Promise<ExcelFileHandle[]>
    }).showOpenFilePicker
    if (!picker) {
      fileInputRef.current?.click()
      return
    }
    const [handle] = await picker({
      multiple: false,
      types: [{ description: 'Excel workbook', accept: { 'application/vnd.ms-excel.sheet.macroEnabled.12': ['.xlsm'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
    })
    const file = await handle.getFile()
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', bookVBA: true })
    applyWorkbook(workbook, file.name)
    setExcelFileHandle(handle)
  }

  async function saveWorkbook(eventsToSave = events) {
    const workbook = workbookState?.workbook || XLSX.utils.book_new()
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    if (sheet) {
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, defval: null, raw: true })
      const headerIndex = rows.findIndex((row) => row.some((cell) => String(cell || '').includes('תאריך לועזי')))
      if (headerIndex >= 0) {
        const headers = rows[headerIndex].map((cell) => String(cell || ''))
        const dateIndex = headers.findIndex((header) => header.includes('תאריך לועזי'))
        const titleIndex = headers.findIndex((header) => header.includes('שם החוגגת'))
        const classIndex = headers.findIndex((header) => header.includes('כיתה'))
        if (dateIndex >= 0) {
          const columns = sheet['!cols'] || []
          columns[dateIndex] = { ...(columns[dateIndex] || {}), wch: 14 }
          sheet['!cols'] = columns
        }
        const eventsByDate = new Map(eventsToSave.map((item) => [item.date, item]))
        rows.slice(headerIndex + 1).forEach((row, index) => {
          const date = dateIndex >= 0 ? normalizeExcelDate(row[dateIndex]) : ''
          const rowNumber = headerIndex + 1 + index
          if (dateIndex >= 0 && date) {
            const address = XLSX.utils.encode_cell({ r: rowNumber, c: dateIndex })
            sheet[address] = { ...(sheet[address] || {}), t: 'd', v: excelDateFromDateKey(date), z: 'dd/mm/yyyy' }
          }
          const item = eventsByDate.get(date)
          if (!item) return
          if (titleIndex >= 0) {
            const address = XLSX.utils.encode_cell({ r: rowNumber, c: titleIndex })
            sheet[address] = { ...(sheet[address] || {}), t: 's', v: item.title }
          }
          if (classIndex >= 0) {
            const address = XLSX.utils.encode_cell({ r: rowNumber, c: classIndex })
            sheet[address] = { ...(sheet[address] || {}), t: 's', v: item.className }
          }
        })
      }
    } else {
      const rows = eventsToSave.map((item) => ({ 'תאריך לועזי': excelDateFromDateKey(item.date), 'שם החוגגת': item.title, 'כיתה': item.className }))
      const newSheet = XLSX.utils.json_to_sheet(rows)
      newSheet['!cols'] = [{ wch: 14 }, { wch: 24 }, { wch: 12 }]
      eventsToSave.forEach((item, index) => {
        const address = XLSX.utils.encode_cell({ r: index + 1, c: 0 })
        newSheet[address] = { ...(newSheet[address] || {}), t: 'd', v: excelDateFromDateKey(item.date), z: 'dd/mm/yyyy' }
      })
      XLSX.utils.book_append_sheet(workbook, newSheet, 'אירועים')
    }
    const sourceExtension = workbookState?.fileName.match(/\.(xlsx|xlsm|xls)$/i)?.[1].toLowerCase() || 'xlsm'
    const fileName = workbookState?.fileName.replace(/\.(xlsx|xlsm|xls)$/i, '') || 'תאריכיבתמצוות'
    const workbookData = XLSX.write(workbook, { bookType: sourceExtension as 'xlsx' | 'xlsm' | 'xls', type: 'array', bookVBA: sourceExtension === 'xlsm', cellDates: true }) as ArrayBuffer
    if (excelFileHandle) {
      const writable = await excelFileHandle.createWritable()
      await writable.write(workbookData)
      await writable.close()
    } else {
      XLSX.writeFile(workbook, `${fileName}.${sourceExtension}`, { bookType: sourceExtension as 'xlsx' | 'xlsm' | 'xls', bookVBA: sourceExtension === 'xlsm', cellDates: true })
    }
    setWorkbookState({ workbook, fileName: `${fileName}.${sourceExtension}` })
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-block">
          <button
            className={`eyebrow ${isAdmin ? 'admin-active' : ''}`}
            type="button"
            onClick={requestAdminAccess}
            title={isAdmin ? 'יציאה ממצב מנהל' : 'כניסה למצב מנהל'}
          >Carmit Vaknin Software</button>
          <h1>לוח אירועים שנת בת מצווה</h1>
        </div>
        <button className="primary-button" onClick={() => { setSelectedEventDate(''); setIsFormOpen(true) }}>
          <span className="plus-icon">+</span>
          שבץ אירוע
        </button>
      </header>

      {isPasswordDialogOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setIsPasswordDialogOpen(false)}>
          <section className="password-card" role="dialog" aria-modal="true" aria-labelledby="password-title">
            <button className="close-button" type="button" onClick={() => setIsPasswordDialogOpen(false)} aria-label="סגירת חלון הסיסמה">×</button>
            <h2 id="password-title">הכנס סיסמה</h2>
            <form onSubmit={verifyAdminPassword}>
              <label>
                סיסמה
                <span className="password-input-wrap">
                  <input
                    type={showAdminPassword ? 'text' : 'password'}
                    value={adminPassword}
                    autoFocus
                    onChange={(event) => { setAdminPassword(event.target.value); setPasswordError(false) }}
                    required
                  />
                  <button
                    type="button"
                    className="password-visibility-button"
                    onClick={() => setShowAdminPassword((current) => !current)}
                    aria-label={showAdminPassword ? 'הסתרת סיסמה' : 'הצגת סיסמה'}
                    title={showAdminPassword ? 'הסתרת סיסמה' : 'הצגת סיסמה'}
                  >👁</button>
                </span>
              </label>
              {passwordError && <span className="field-error">סיסמה שגויה</span>}
              <button className="submit-button" type="submit">אישור</button>
            </form>
          </section>
        </div>
      )}

      {deleteConfirmDate && (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setDeleteConfirmDate('')}>
          <section className="confirmation-card" role="dialog" aria-modal="true" aria-labelledby="delete-title">
            <button className="close-button" type="button" onClick={() => setDeleteConfirmDate('')} aria-label="סגירת חלון האישור">×</button>
            <div className="confirmation-icon">!</div>
            <h2 id="delete-title">מחיקת אירוע</h2>
            <p>האם למחוק את האירוע מהלוח ומהאקסל?</p>
            <div className="confirmation-actions">
              <button type="button" className="cancel-button" onClick={() => setDeleteConfirmDate('')}>ביטול</button>
              <button type="button" className="delete-confirm-button" onClick={() => void deleteEvent(deleteConfirmDate)}>מחק</button>
            </div>
          </section>
        </div>
      )}

      <section className="calendar-panel" aria-label="לוח שנה">
        <div className="calendar-heading">
          <button className="month-arrow" onClick={() => moveMonth(1)} aria-label="חודש הבא">→</button>
          <div>
            <p className="month-kicker">{getHebrewYearRange(visibleMonth)} · {visibleMonth.getFullYear()}</p>
            <h2>{monthNames[visibleMonth.getMonth()]}</h2>
          </div>
          <button
            className="month-arrow"
            onClick={() => moveMonth(-1)}
            disabled={visibleMonth.getFullYear() === today.getFullYear() && visibleMonth.getMonth() === today.getMonth()}
            aria-label="חודש קודם"
          >←</button>
        </div>

        <div className="footer-note calendar-legend">
          <span className="legend-item"><span className="legend-dot available-dot" />פנוי</span>
          <span className="legend-item"><span className="legend-dot occupied-dot" />תפוס</span>
          <span className="excel-actions">
            <input ref={fileInputRef} className="visually-hidden" type="file" accept=".xlsm,.xlsx" onChange={loadWorkbook} />
            <button type="button" className="excel-button" onClick={() => void saveWorkbook()}>הורד נתונים כ Excel</button>
          </span>
        </div>

        <div className="week-row">
          {dayNames.map((day) => <span key={day}>{day}</span>)}
        </div>
        <div className="calendar-grid">
          {calendarDays.map((date, index) => {
            const dateKey = date ? toDateKey(date) : `empty-${index}`
            const dayEvents = date ? events.filter((item) => item.date === dateKey) : []
            const isToday = dateKey === toDateKey(today)
            const isOccupied = dayEvents.length > 0
            return (
              <div
                className={`day-cell ${!date ? 'empty' : isOccupied ? 'occupied' : 'available'} ${isToday ? 'today' : ''}`}
                key={dateKey}
                onClick={() => isAdmin && isOccupied && setDeleteCandidateDate(dateKey)}
                onContextMenu={(event) => {
                  if (!isAdmin || !isOccupied) return
                  event.preventDefault()
                  setDeleteCandidateDate(dateKey)
                }}
              >
                {date && (
                  <div className="day-label">
                    <span className={`day-number ${isOccupied ? 'occupied' : 'available'}`}>{date.getDate()}</span>
                    <span className="hebrew-date">{getHebrewDate(date)}</span>
                  </div>
                )}
                {dayEvents.map((item) => (
                  <span
                    className="event-pill"
                    key={item.id}
                    onDoubleClick={(event) => { event.stopPropagation(); setSelectedEvent(item) }}
                    title="לחצי פעמיים להצגת פרטי האירוע"
                  >{item.title}</span>
                ))}
                {isAdmin && isOccupied && deleteCandidateDate === dateKey && (
                  <button
                    type="button"
                    className="delete-event-button"
                    onClick={(event) => { event.stopPropagation(); setDeleteConfirmDate(dateKey) }}
                  >מחק</button>
                )}
              </div>
            )
          })}
        </div>
      </section>

      {selectedEvent && (
        <div className="modal-backdrop event-details-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setSelectedEvent(null)}>
          <section className="event-details-card" role="dialog" aria-modal="true" aria-labelledby="event-details-title">
            <button className="close-button event-details-close-button" type="button" onClick={() => setSelectedEvent(null)} aria-label="סגירת פרטי האירוע">×</button>
            <div className="event-card-heading">
              <span className="mini-spark">✦</span>
              <h2 id="event-details-title">פרטי האירוע</h2>
            </div>
            <dl className="event-details-list">
              <div><dt>שם האירוע</dt><dd>{selectedEvent.title}</dd></div>
              <div><dt>כיתה</dt><dd>{selectedEvent.className}</dd></div>
              <div><dt>תאריך האירוע</dt><dd>{formatDateForDisplay(selectedEvent.date)}</dd></div>
            </dl>
          </section>
        </div>
      )}

      {isFormOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={(event) => event.target === event.currentTarget && setIsFormOpen(false)}>
          <section className="event-card" role="dialog" aria-modal="true" aria-labelledby="event-title">
            <button className="close-button" onClick={() => setIsFormOpen(false)} aria-label="סגירת הטופס">×</button>
            <div className="event-card-heading">
              <span className="mini-spark">✦</span>
              <h2 id="event-title">שיבוץ אירוע</h2>
              <p>שיהיה תמיד בשמחות!</p>
            </div>
            <form onSubmit={addEvent} noValidate>
              <label>
                שם האירוע
                <input
                  name="title"
                  placeholder="לדוגמא: בת מצווה ל..  טיול שנתי,  מסיבת סיום"
                  required
                  onChange={() => setValidationErrors((current) => ({ ...current, title: false }))}
                />
                {validationErrors.title && <span className="field-error">שדה חובה למילוי</span>}
              </label>
              <label>
                כיתה
                <select
                  name="className"
                  defaultValue=""
                  required
                  onChange={() => setValidationErrors((current) => ({ ...current, className: false }))}
                >
                  <option value="" disabled>בחר כיתה</option>
                  <option value="ו3">ו3</option>
                  <option value="ו4">ו4</option>
                  <option value="שכבתי">שכבתי</option>
                  <option value="בית ספרי">בית ספרי</option>
                </select>
                {validationErrors.className && <span className="field-error">שדה חובה למילוי</span>}
              </label>
              <label>
                תאריך האירוע
                <EventDatePicker
                  selectedDate={selectedEventDate}
                  onSelect={setSelectedEventDate}
                  events={events}
                  showError={dateValidationAttempted && !selectedEventDate}
                />
              </label>
              <button className="submit-button" type="submit">שבץ אירוע</button>
            </form>
          </section>
        </div>
      )}
    </main>
  )
}

export default App