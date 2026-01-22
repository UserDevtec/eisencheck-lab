import { useEffect, useMemo, useRef, useState } from 'react'
import ExcelJS from 'exceljs'
import './App.css'

const STATUS_UNCHANGED = 'Ongewijzigd'
const STATUS_ADDED = 'Toegevoegd'
const STATUS_CHANGED = 'Gewijzigd'

const LEGEND_ITEMS = [
  { status: STATUS_UNCHANGED, color: 'green', note: 'Niets veranderd' },
  { status: STATUS_ADDED, color: 'yellow', note: 'Nieuw toegevoegd in bestand B' },
  { status: STATUS_CHANGED, color: 'orange', note: 'Waarde gewijzigd t.o.v. bestand A' },
]

const normalizeVisible = (value) =>
  String(value ?? '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim()

const getCellText = (cell) => {
  if (!cell) return ''
  if (cell.text != null && String(cell.text).trim() !== '') return String(cell.text)
  const value = cell.value
  if (value == null) return ''
  if (typeof value === 'object' && value.richText) {
    return value.richText.map((part) => part.text).join('')
  }
  if (value instanceof Date) return value.toISOString()
  return String(value)
}

const parseSheet = (sheet) => {
  const headerRow = sheet.getRow(1)
  const maxColumns = Math.max(headerRow.cellCount, headerRow.actualCellCount)
  const headerValues = []
  for (let col = 1; col <= maxColumns; col += 1) {
    headerValues[col - 1] = getCellText(headerRow.getCell(col))
  }
  while (headerValues.length && normalizeVisible(headerValues[headerValues.length - 1]) === '') {
    headerValues.pop()
  }
  const headers = headerValues.map((value) => String(value ?? ''))
  const rows = []
  for (let rowIndex = 2; rowIndex <= sheet.rowCount; rowIndex += 1) {
    const row = sheet.getRow(rowIndex)
    const values = []
    let hasContent = false
    for (let col = 1; col <= headers.length; col += 1) {
      const text = getCellText(row.getCell(col))
      values.push(text)
      if (normalizeVisible(text) !== '') {
        hasContent = true
      }
    }
    if (hasContent) rows.push(values)
  }
  return { headers, rows }
}

const buildHeaderCounts = (headers) => {
  const counts = new Map()
  headers
    .map((header) => normalizeVisible(header))
    .filter((header) => header !== '')
    .forEach((header) => {
      counts.set(header, (counts.get(header) ?? 0) + 1)
    })
  return counts
}

const compareHeaders = (headersA, headersB) => {
  const countsA = buildHeaderCounts(headersA)
  const countsB = buildHeaderCounts(headersB)
  const missingInA = []
  const missingInB = []
  for (const [header, countB] of countsB.entries()) {
    const countA = countsA.get(header) ?? 0
    if (countA < countB) missingInA.push(header)
  }
  for (const [header, countA] of countsA.entries()) {
    const countB = countsB.get(header) ?? 0
    if (countB < countA) missingInB.push(header)
  }
  return {
    ok: missingInA.length === 0 && missingInB.length === 0,
    missingInA,
    missingInB,
  }
}

const buildIndex = (rows, keyIndex, valueIndex) => {
  const map = new Map()
  const keyOrder = []
  const duplicates = new Set()
  let emptyKeys = 0
  rows.forEach((row, idx) => {
    const rawKey = row[keyIndex] ?? ''
    const normKey = normalizeVisible(rawKey)
    if (!normKey) {
      emptyKeys += 1
      return
    }
    if (!map.has(normKey)) {
      map.set(normKey, [])
      keyOrder.push(normKey)
    }
    const list = map.get(normKey)
    if (list.length) duplicates.add(normKey)
    list.push({
      rawKey: String(rawKey ?? ''),
      rawValue: String(row[valueIndex] ?? ''),
      normValue: normalizeVisible(row[valueIndex] ?? ''),
      rowIndex: idx + 2,
    })
  })
  return { map, keyOrder, duplicates, emptyKeys }
}

const matchByValue = (listA, listB) => {
  const queuesA = new Map()
  listA.forEach((item) => {
    const queue = queuesA.get(item.normValue) ?? []
    queue.push(item)
    queuesA.set(item.normValue, queue)
  })
  const matched = []
  const remainingB = []
  listB.forEach((item) => {
    const queue = queuesA.get(item.normValue)
    if (queue && queue.length) {
      matched.push({ a: queue.shift(), b: item })
    } else {
      remainingB.push(item)
    }
  })
  const remainingA = []
  for (const queue of queuesA.values()) {
    remainingA.push(...queue)
  }
  return { matched, remainingA, remainingB }
}

const pickIndex = (current, headers, fallback) => {
  if (!headers.length) return ''
  const maxIndex = headers.length - 1
  if (current === '' || Number(current) > maxIndex) return String(fallback)
  return current
}

function App() {
  const [dataA, setDataA] = useState(null)
  const [dataB, setDataB] = useState(null)
  const [error, setError] = useState('')
  const [keyColA, setKeyColA] = useState('')
  const [keyColB, setKeyColB] = useState('')
  const [compareColA, setCompareColA] = useState('')
  const [compareColB, setCompareColB] = useState('')
  const [results, setResults] = useState(null)
  const [activeTab, setActiveTab] = useState('result')
  const [draggingA, setDraggingA] = useState(false)
  const [draggingB, setDraggingB] = useState(false)
  const [showHelp, setShowHelp] = useState(false)
  const [helpItems, setHelpItems] = useState([])
  const [isHelpCompact, setIsHelpCompact] = useState(false)
  const helpRafRef = useRef(null)
  const uploadsRef = useRef(null)
  const columnsRef = useRef(null)
  const compareRef = useRef(null)
  const resultsRef = useRef(null)

  const headerCheck = useMemo(() => {
    if (!dataA || !dataB) return null
    return compareHeaders(dataA.headers, dataB.headers)
  }, [dataA, dataB])

  useEffect(() => {
    if (!showHelp) return
    const buildHelpItems = () => {
      const compact = window.innerWidth <= 1024 || window.innerHeight <= 720
      setIsHelpCompact(compact)
      const padding = 16
      const cardWidth = 280
      const cardGap = 14
      const clamp = (value, min, max) => Math.min(Math.max(value, min), max)
      const withRect = (ref) => {
        if (!ref.current) return null
        return ref.current.getBoundingClientRect()
      }
      const placeCard = (rect, preferLeft = false) => {
        const roomRight = window.innerWidth - rect.right
        const roomLeft = rect.left
        const canPlaceRight = roomRight >= cardWidth + cardGap
        const canPlaceLeft = roomLeft >= cardWidth + cardGap
        let left = clamp(rect.left, padding, window.innerWidth - cardWidth - padding)
        if (preferLeft) {
          if (canPlaceLeft) {
            left = rect.left - cardWidth - cardGap
          } else if (canPlaceRight) {
            left = rect.right + cardGap
          }
        } else if (canPlaceRight) {
          left = rect.right + cardGap
        }
        const top = clamp(rect.top, padding, window.innerHeight - 140)
        return { left, top }
      }
      const nextItems = []
      const uploadsRect = withRect(uploadsRef)
      if (uploadsRect) {
        nextItems.push({
          id: 'uploads',
          title: 'Stap 1: Upload',
          body: 'Upload het oude en nieuwe Excel-bestand (.xlsx).',
          ...(compact
            ? {}
            : {
                spot: {
                  top: uploadsRect.top,
                  left: uploadsRect.left,
                  width: uploadsRect.width,
                  height: uploadsRect.height,
                },
                card: placeCard(uploadsRect),
              }),
        })
      }
      const columnsRect = withRect(columnsRef)
      if (columnsRect) {
        nextItems.push({
          id: 'columns',
          title: 'Stap 2: Kolommen',
          body: 'Kies de sleutelkolommen en de kolom met eisen-tekst.',
          ...(compact
            ? {}
            : {
                spot: {
                  top: columnsRect.top,
                  left: columnsRect.left,
                  width: columnsRect.width,
                  height: columnsRect.height,
                },
                card: placeCard(columnsRect),
              }),
        })
      }
      const compareRect = withRect(compareRef)
      if (compareRect) {
        nextItems.push({
          id: 'actions',
          title: 'Stap 3: Vergelijk',
          body: 'Klik "Vergelijk bestanden" in deze kaart.',
          ...(compact
            ? {}
            : {
                spot: {
                  top: compareRect.top,
                  left: compareRect.left,
                  width: compareRect.width,
                  height: compareRect.height,
                },
                card: placeCard(compareRect, true),
              }),
        })
      }
      const resultsRect = withRect(resultsRef)
      if (resultsRect) {
        nextItems.push({
          id: 'results',
          title: 'Stap 4: Resultaten',
          body: 'Bekijk de gemarkeerde verschillen en download het Excel-overzicht.',
          ...(compact
            ? {}
            : {
                spot: {
                  top: resultsRect.top,
                  left: resultsRect.left,
                  width: resultsRect.width,
                  height: resultsRect.height,
                },
                card: placeCard(resultsRect),
              }),
        })
      }
      setHelpItems(nextItems)
    }
    const scheduleUpdate = () => {
      if (helpRafRef.current != null) return
      helpRafRef.current = requestAnimationFrame(() => {
        helpRafRef.current = null
        buildHelpItems()
      })
    }
    scheduleUpdate()
    window.addEventListener('resize', scheduleUpdate)
    window.addEventListener('scroll', scheduleUpdate, true)
    return () => {
      window.removeEventListener('resize', scheduleUpdate)
      window.removeEventListener('scroll', scheduleUpdate, true)
      if (helpRafRef.current != null) {
        cancelAnimationFrame(helpRafRef.current)
        helpRafRef.current = null
      }
    }
  }, [showHelp])

  const loadFile = async (file, side) => {
    if (!file) return
    setError('')
    setResults(null)
    try {
      const workbook = new ExcelJS.Workbook()
      const buffer = await file.arrayBuffer()
      await workbook.xlsx.load(buffer)
      const sheet = workbook.worksheets[0]
      if (!sheet) throw new Error('Geen werkblad gevonden in het bestand.')
      const parsed = parseSheet(sheet)
      if (!parsed.headers.length) {
        throw new Error('Geen headers gevonden in de eerste rij.')
      }
      const payload = {
        fileName: file.name,
        sheetName: sheet.name || 'Werkblad 1',
        headers: parsed.headers,
        rows: parsed.rows,
      }
      if (side === 'A') {
        setDataA(payload)
        setKeyColA((prev) => pickIndex(prev, payload.headers, 0))
        setCompareColA((prev) => pickIndex(prev, payload.headers, payload.headers.length > 1 ? 1 : 0))
      } else {
        setDataB(payload)
        setKeyColB((prev) => pickIndex(prev, payload.headers, 0))
        setCompareColB((prev) => pickIndex(prev, payload.headers, payload.headers.length > 1 ? 1 : 0))
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err)
      setError(message)
    }
  }

  const canCompare =
    dataA &&
    dataB &&
    headerCheck?.ok &&
    keyColA !== '' &&
    keyColB !== '' &&
    compareColA !== '' &&
    compareColB !== ''

  const oldHeaderLabel = dataA
    ? `${dataA.headers[Number(compareColA)] || 'EisTekst'} [Oud]`
    : 'EisTekst [Oud]'
  const newHeaderLabel = dataB
    ? `${dataB.headers[Number(compareColB)] || 'EisTekst'} [Nieuw]`
    : 'EisTekst [Nieuw]'

  const runCompare = () => {
    if (!dataA || !dataB) return
    if (!headerCheck?.ok) return
    setError('')
    const keyIndexA = Number(keyColA)
    const keyIndexB = Number(keyColB)
    const valueIndexA = Number(compareColA)
    const valueIndexB = Number(compareColB)
    if ([keyIndexA, keyIndexB, valueIndexA, valueIndexB].some((val) => Number.isNaN(val))) {
      setError('Selecteer geldige kolommen om te vergelijken.')
      return
    }
    const indexA = buildIndex(dataA.rows, keyIndexA, valueIndexA)
    const indexB = buildIndex(dataB.rows, keyIndexB, valueIndexB)

    const rows = []
    const removed = []
    let unchangedCount = 0
    let addedCount = 0
    let changedCount = 0

    const keysFromB = indexB.keyOrder
    const keysFromA = indexA.keyOrder.filter((key) => !indexB.map.has(key))
    const allKeys = [...keysFromB, ...keysFromA]

    allKeys.forEach((key) => {
      const listA = indexA.map.get(key) ?? []
      const listB = indexB.map.get(key) ?? []
      if (!listA.length && listB.length) {
        listB.forEach((item) => {
          rows.push({
            status: STATUS_ADDED,
            key: item.rawKey,
            oldValue: '',
            newValue: item.rawValue,
          })
          addedCount += 1
        })
        return
      }
      if (!listB.length && listA.length) {
        listA.forEach((item) => {
          removed.push({
            key: item.rawKey,
            oldValue: item.rawValue,
          })
        })
        return
      }

      const matched = matchByValue(listA, listB)
      matched.matched.forEach(({ a, b }) => {
        rows.push({
          status: STATUS_UNCHANGED,
          key: b.rawKey || a.rawKey,
          oldValue: a.rawValue,
          newValue: b.rawValue,
        })
        unchangedCount += 1
      })

      const remainingA = [...matched.remainingA]
      const remainingB = [...matched.remainingB]
      while (remainingA.length && remainingB.length) {
        const a = remainingA.shift()
        const b = remainingB.shift()
        rows.push({
          status: STATUS_CHANGED,
          key: b.rawKey || a.rawKey,
          oldValue: a.rawValue,
          newValue: b.rawValue,
        })
        changedCount += 1
      }
      remainingB.forEach((item) => {
        rows.push({
          status: STATUS_ADDED,
          key: item.rawKey,
          oldValue: '',
          newValue: item.rawValue,
        })
        addedCount += 1
      })
      remainingA.forEach((item) => {
        removed.push({
          key: item.rawKey,
          oldValue: item.rawValue,
        })
      })
    })

    setResults({
      rows,
      removed,
      stats: {
        unchanged: unchangedCount,
        added: addedCount,
        changed: changedCount,
        removed: removed.length,
      },
      duplicatesA: indexA.duplicates,
      duplicatesB: indexB.duplicates,
      emptyKeysA: indexA.emptyKeys,
      emptyKeysB: indexB.emptyKeys,
    })
    setActiveTab('result')
  }

  const downloadExcel = async () => {
    if (!results || !dataA || !dataB) return
    const workbook = new ExcelJS.Workbook()
    workbook.creator = 'Eisencheck Lab'
    workbook.created = new Date()

    const resultSheet = workbook.addWorksheet('Resultaat')
    const resultHeaders = ['Eiscode', oldHeaderLabel, newHeaderLabel, 'Status']
    resultSheet.addRow(resultHeaders)
    results.rows.forEach((row) => {
      const excelRow = resultSheet.addRow([row.key, row.oldValue, row.newValue, row.status])
      let fillColor = null
      if (row.status === STATUS_UNCHANGED) fillColor = 'FFE3F5DF'
      if (row.status === STATUS_ADDED) fillColor = 'FFFFF2C5'
      if (row.status === STATUS_CHANGED) fillColor = 'FFFFE6B7'
      if (fillColor) {
        excelRow.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: fillColor },
          }
        })
      }
    })

    const removedSheet = workbook.addWorksheet('Vervallen eisen')
    removedSheet.addRow(['Eiscode', oldHeaderLabel])
    results.removed.forEach((row) => {
      const excelRow = removedSheet.addRow([row.key, row.oldValue])
      excelRow.eachCell((cell) => {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFD9D1' },
        }
      })
    })

    const legendSheet = workbook.addWorksheet('Legenda')
    legendSheet.addRow(['Status', 'Betekenis', 'Kleur'])
    LEGEND_ITEMS.forEach((item) => {
      legendSheet.addRow([item.status, item.note, item.color])
    })
    legendSheet.addRow(['Vervallen', 'Alleen in bestand A', 'red'])

    const sheets = [resultSheet, removedSheet, legendSheet]
    sheets.forEach((sheet) => {
      sheet.getRow(1).font = { bold: true }
      sheet.columns = sheet.columns.map((col) => ({
        ...col,
        width: Math.min(Math.max(col.width || 10, 16), 60),
      }))
      sheet.eachRow((row) => {
        row.alignment = { vertical: 'top', wrapText: true }
      })
    })

    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')
    link.href = url
    link.download = `eisencheck_${new Date().toISOString().slice(0, 10)}.xlsx`
    link.rel = 'noopener'
    link.click()
    setTimeout(() => URL.revokeObjectURL(url), 1000)
  }

  const renderRows = (rowsToRender) => {
    if (!rowsToRender.length) {
      return <p className="note">Geen rijen om te tonen.</p>
    }
    return (
      <div className="compare-table">
        <table>
          <thead>
            <tr>
              <th>Eiscode</th>
              <th>{oldHeaderLabel}</th>
              <th>{newHeaderLabel}</th>
              <th>Status</th>
            </tr>
          </thead>
          <tbody>
            {rowsToRender.map((row, index) => {
              const statusClass =
                row.status === STATUS_UNCHANGED
                  ? 'status-unchanged'
                  : row.status === STATUS_ADDED
                    ? 'status-added'
                    : 'status-changed'
              return (
                <tr key={`${row.key}-${index}`} className={statusClass}>
                  <td>{row.key}</td>
                  <td>{row.oldValue}</td>
                  <td>{row.newValue}</td>
                  <td>{row.status}</td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    )
  }

  const renderRemoved = (rowsToRender) => {
    if (!rowsToRender.length) {
      return <p className="note">Geen vervallen eisen.</p>
    }
    return (
      <div className="compare-table">
        <table>
          <thead>
            <tr>
              <th>Eiscode</th>
              <th>{oldHeaderLabel}</th>
            </tr>
          </thead>
          <tbody>
            {rowsToRender.map((row, index) => (
              <tr key={`${row.key}-${index}`} className="status-removed">
                <td>{row.key}</td>
                <td>{row.oldValue}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    )
  }

  return (
    <div className="app">
      <header className="hero">
        <div>
          <p className="eyebrow">Eisencheck Lab</p>
          <h1>Excel vergelijken in dezelfde stijl</h1>
          <p className="subtitle">
            Upload twee Excel-bestanden, kies de sleutelkolom en de kolom om te
            vergelijken. De app negeert volgorde en normaliseert onzichtbare
            tekens zodat alleen zichtbare verschillen tellen.
          </p>
        </div>
      </header>

      <section className="panel" ref={uploadsRef}>
        <div className="panel-header">
          <h2>Upload bestanden</h2>
        </div>
        <div className="upload-grid">
          <div className="upload-card">
            <strong>Bestand A (oud)</strong>
            <div
              className={`upload ${draggingA ? 'dragging' : ''}`}
              onDragOver={(event) => {
                event.preventDefault()
                setDraggingA(true)
              }}
              onDragLeave={() => setDraggingA(false)}
              onDrop={(event) => {
                event.preventDefault()
                setDraggingA(false)
                const file = event.dataTransfer.files?.[0]
                loadFile(file, 'A')
              }}
            >
              <label className="file">
                Kies bestand
                <input
                  type="file"
                  accept=".xlsx"
                  onChange={(event) => loadFile(event.target.files?.[0], 'A')}
                />
              </label>
              <span className="hint">Sleep of klik om een .xlsx te kiezen.</span>
            </div>
            {dataA ? (
              <div className="info-row">
                <span>
                  <strong>Naam:</strong> {dataA.fileName}
                </span>
                <span>
                  <strong>Sheet:</strong> {dataA.sheetName}
                </span>
                <span>
                  <strong>Rijen:</strong> {dataA.rows.length}
                </span>
              </div>
            ) : (
              <p className="upload-status">Nog geen bestand geladen.</p>
            )}
          </div>

          <div className="upload-card">
            <strong>Bestand B (nieuw)</strong>
            <div
              className={`upload ${draggingB ? 'dragging' : ''}`}
              onDragOver={(event) => {
                event.preventDefault()
                setDraggingB(true)
              }}
              onDragLeave={() => setDraggingB(false)}
              onDrop={(event) => {
                event.preventDefault()
                setDraggingB(false)
                const file = event.dataTransfer.files?.[0]
                loadFile(file, 'B')
              }}
            >
              <label className="file">
                Kies bestand
                <input
                  type="file"
                  accept=".xlsx"
                  onChange={(event) => loadFile(event.target.files?.[0], 'B')}
                />
              </label>
              <span className="hint">Sleep of klik om een .xlsx te kiezen.</span>
            </div>
            {dataB ? (
              <div className="info-row">
                <span>
                  <strong>Naam:</strong> {dataB.fileName}
                </span>
                <span>
                  <strong>Sheet:</strong> {dataB.sheetName}
                </span>
                <span>
                  <strong>Rijen:</strong> {dataB.rows.length}
                </span>
              </div>
            ) : (
              <p className="upload-status">Nog geen bestand geladen.</p>
            )}
          </div>
        </div>
        {error ? <div className="error">{error}</div> : null}
      </section>

      <section className="panel" ref={columnsRef}>
        <div className="panel-header">
          <h2>Kolommen kiezen</h2>
        </div>
        <div className="select-grid">
          <div className="select-field">
            <label htmlFor="keyA">Eiscode kolom (A)</label>
            <select
              id="keyA"
              value={keyColA}
              onChange={(event) => setKeyColA(event.target.value)}
              disabled={!dataA}
            >
              <option value="" disabled>
                Kies kolom
              </option>
              {dataA?.headers.map((header, index) => (
                <option key={`a-key-${header}-${index}`} value={String(index)}>
                  {header || `(kolom ${index + 1})`}
                </option>
              ))}
            </select>
          </div>
          <div className="select-field">
            <label htmlFor="keyB">Eiscode kolom (B)</label>
            <select
              id="keyB"
              value={keyColB}
              onChange={(event) => setKeyColB(event.target.value)}
              disabled={!dataB}
            >
              <option value="" disabled>
                Kies kolom
              </option>
              {dataB?.headers.map((header, index) => (
                <option key={`b-key-${header}-${index}`} value={String(index)}>
                  {header || `(kolom ${index + 1})`}
                </option>
              ))}
            </select>
          </div>
          <div className="select-field">
            <label htmlFor="compareA">Vergelijk kolom (A)</label>
            <select
              id="compareA"
              value={compareColA}
              onChange={(event) => setCompareColA(event.target.value)}
              disabled={!dataA}
            >
              <option value="" disabled>
                Kies kolom
              </option>
              {dataA?.headers.map((header, index) => (
                <option key={`a-comp-${header}-${index}`} value={String(index)}>
                  {header || `(kolom ${index + 1})`}
                </option>
              ))}
            </select>
          </div>
          <div className="select-field">
            <label htmlFor="compareB">Vergelijk kolom (B)</label>
            <select
              id="compareB"
              value={compareColB}
              onChange={(event) => setCompareColB(event.target.value)}
              disabled={!dataB}
            >
              <option value="" disabled>
                Kies kolom
              </option>
              {dataB?.headers.map((header, index) => (
                <option key={`b-comp-${header}-${index}`} value={String(index)}>
                  {header || `(kolom ${index + 1})`}
                </option>
              ))}
            </select>
          </div>
        </div>
        {headerCheck && !headerCheck.ok ? (
          <div className="error">
            Headers komen niet overeen.
            {headerCheck.missingInA.length ? (
              <div>Ontbreekt in A: {headerCheck.missingInA.join(', ')}</div>
            ) : null}
            {headerCheck.missingInB.length ? (
              <div>Ontbreekt in B: {headerCheck.missingInB.join(', ')}</div>
            ) : null}
          </div>
        ) : null}
        <p className="note">
          Vergelijking negeert dubbele spaties, returns en onzichtbare tekens.
        </p>
      </section>

      <section className="panel compare-panel" ref={compareRef}>
        <div className="panel-header">
          <h2>Vergelijk geuploade bestanden</h2>
        </div>
        <div className="upload-actions">
          <button className="primary" type="button" onClick={runCompare} disabled={!canCompare}>
            Vergelijk bestanden
          </button>
        </div>
      </section>

      <section className="panel results" ref={resultsRef}>
        <div className="panel-header">
          <div className="panel-title-row">
            <h2>Resultaten</h2>
            <button className="ghost" type="button" onClick={downloadExcel} disabled={!results}>
              Download Excel
            </button>
          </div>
          <div className="panel-actions">
            <span className="pill green">Groen: niets veranderd</span>
            <span className="pill yellow">Geel: toegevoegd</span>
            <span className="pill orange">Oranje: gewijzigd</span>
            <span className="pill red">Rood: vervallen</span>
          </div>
        </div>

        <div className="output-cards">
          <div className="stat-card">
            <p className="stat-label">Ongewijzigd</p>
            <p className="stat-value">{results?.stats.unchanged ?? 0}</p>
            <p className="stat-note">Bestanden matchen exact</p>
          </div>
          <div className="stat-card">
            <p className="stat-label">Toegevoegd</p>
            <p className="stat-value">{results?.stats.added ?? 0}</p>
            <p className="stat-note">Alleen in bestand B</p>
          </div>
          <div className="stat-card">
            <p className="stat-label">Gewijzigd</p>
            <p className="stat-value">{results?.stats.changed ?? 0}</p>
            <p className="stat-note">Zelfde sleutel, andere waarde</p>
          </div>
          <div className="stat-card">
            <p className="stat-label">Vervallen eisen</p>
            <p className="stat-value">{results?.stats.removed ?? 0}</p>
            <p className="stat-note">Alleen in bestand A</p>
          </div>
        </div>

        {results?.duplicatesA?.size || results?.duplicatesB?.size ? (
          <div className="warning">
            Let op: dubbele sleutels gevonden. Matches worden op tekstwaarde gepaard.
          </div>
        ) : null}
        {(results?.emptyKeysA || results?.emptyKeysB) && results ? (
          <div className="warning">
            Lege sleutelwaarden genegeerd (A: {results.emptyKeysA}, B: {results.emptyKeysB}).
          </div>
        ) : null}

        <div className="legend">
          {LEGEND_ITEMS.map((item) => (
            <div key={item.status} className="legend-item">
              <span className={`legend-swatch ${item.color}`} />
              <div>
                <strong>{item.status}</strong>
                <div>{item.note}</div>
              </div>
            </div>
          ))}
          <div className="legend-item">
            <span className="legend-swatch red" />
            <div>
              <strong>Vervallen</strong>
              <div>Verplaatst naar apart tabblad</div>
            </div>
          </div>
        </div>

        <div className="tab-row">
          <button
            className={`tab-button ${activeTab === 'result' ? 'active' : ''}`}
            type="button"
            onClick={() => setActiveTab('result')}
          >
            Resultaat
          </button>
          <button
            className={`tab-button ${activeTab === 'removed' ? 'active' : ''}`}
            type="button"
            onClick={() => setActiveTab('removed')}
          >
            Vervallen eisen
          </button>
        </div>

        <div className="output-section">
          {activeTab === 'result'
            ? renderRows(results?.rows ?? [])
            : renderRemoved(results?.removed ?? [])}
        </div>
      </section>
      <div className={`help-fab-wrap ${showHelp ? 'open' : ''}`}>
        <button className="help-fab" type="button" onClick={() => setShowHelp((prev) => !prev)}>
          ?
        </button>
        <span className="help-fab-label">Hulp nodig? Jonathan van der Gouwe</span>
      </div>
      {showHelp ? (
        <div
          className={`help-overlay ${isHelpCompact ? 'compact' : ''}`}
          onClick={() => {
            setShowHelp(false)
          }}
        >
          <button
            className="help-close ghost"
            type="button"
            onClick={(event) => {
              event.stopPropagation()
              setShowHelp(false)
            }}
          >
            Sluiten
          </button>
          {isHelpCompact ? (
            <div
              className="help-list"
              onClick={(event) => {
                event.stopPropagation()
              }}
            >
              {helpItems.map((item) => (
                <div key={item.id} className="help-card help-card-compact">
                  <h4>{item.title}</h4>
                  <p>{item.body}</p>
                </div>
              ))}
            </div>
          ) : (
            helpItems.map((item) => (
              <div key={item.id} className="help-item">
                <div
                  className="help-spot"
                  style={{
                    top: `${item.spot.top}px`,
                    left: `${item.spot.left}px`,
                    width: `${item.spot.width}px`,
                    height: `${item.spot.height}px`,
                  }}
                />
                <div
                  className="help-card"
                  style={{ top: `${item.card.top}px`, left: `${item.card.left}px` }}
                  onClick={(event) => event.stopPropagation()}
                >
                  <h4>{item.title}</h4>
                  <p>{item.body}</p>
                </div>
              </div>
            ))
          )}
        </div>
      ) : null}
    </div>
  )
}

export default App
