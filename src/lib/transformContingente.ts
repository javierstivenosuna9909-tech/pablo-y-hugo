import type { CellValue, Row, Workbook, Worksheet } from 'exceljs'

export type ContingenteVariant = 'standard' | 'cementos'

type ContingenteLayout = {
  keepSourceCols: readonly number[]
  colorByOutCol: Record<number, string>
  /** Columnas de salida (1-based) a guardar como enteros */
  integerOutCols: readonly number[]
}

const COLOR_APORTE_BENEFICIO = 'D6EAF8'
const COLOR_VALOR_BENEFICIO = 'FCF3CF'
const COLOR_TOTAL_APORTE = 'D5F5E3'
const INTEGER_NUM_FMT = '#,##0'

/** Fragmentos de rótulos conocidos del Contingente (texto normalizado). */
const ROTULO_HINTS = [
  'tipo de entidad',
  'denominacion',
  'id entidad',
  'cuenta banc',
  'clase id',
  'no. documento',
  'no documento',
  'numero de personal',
  'nombre empl',
  'aporte con beneficio tributario',
  'valor del beneficio tributario',
  'aporte sin beneficio',
  'total aporte',
  'tope maximo',
]

/** No Cementos: columnas A/C/H/M/N/P/Q/R fuera */
const LAYOUT_STANDARD: ContingenteLayout = {
  keepSourceCols: [2, 4, 5, 6, 7, 9, 10, 11, 12, 15],
  colorByOutCol: {
    7: COLOR_APORTE_BENEFICIO,
    8: COLOR_VALOR_BENEFICIO,
    10: COLOR_TOTAL_APORTE,
  },
  integerOutCols: [7, 8, 9, 10],
}

/** Cementos: quita B/H/M/N/P/Q/R */
const LAYOUT_CEMENTOS: ContingenteLayout = {
  keepSourceCols: [1, 3, 4, 5, 6, 7, 9, 10, 11, 12, 15],
  colorByOutCol: {
    8: COLOR_APORTE_BENEFICIO,
    9: COLOR_VALOR_BENEFICIO,
    11: COLOR_TOTAL_APORTE,
  },
  integerOutCols: [8, 9, 10, 11],
}

function layoutForVariant(variant: ContingenteVariant): ContingenteLayout {
  return variant === 'cementos' ? LAYOUT_CEMENTOS : LAYOUT_STANDARD
}

export function contingenteVariantFromGrupo(grupoCatalogo: string): ContingenteVariant {
  return grupoCatalogo.trim().toUpperCase() === 'CEMENTOS Y FILIALES'
    ? 'cementos'
    : 'standard'
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ')
}

function cellPlain(value: CellValue): string {
  if (value == null || value === '') return ''
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim()
  }
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object' && 'result' in value) {
    return cellPlain((value as { result: CellValue }).result)
  }
  if (typeof value === 'object' && 'text' in value) {
    return String((value as { text: unknown }).text ?? '').trim()
  }
  return String(value).trim()
}

function isRowEmpty(row: Row, cols: readonly number[]): boolean {
  return cols.every((c) => cellPlain(row.getCell(c).value) === '')
}

/** Fila de rótulos si coincide con varios encabezados conocidos. */
function isRotulosRow(row: Row, cols: readonly number[]): boolean {
  let hits = 0
  for (const c of cols) {
    const n = normalizeText(cellPlain(row.getCell(c).value))
    if (!n) continue
    if (ROTULO_HINTS.some((hint) => n.includes(hint))) hits += 1
  }
  return hits >= 2
}

function toIntegerValue(value: CellValue): number | CellValue {
  if (value == null || value === '') return value
  if (typeof value === 'number' && Number.isFinite(value)) return Math.round(value)
  if (typeof value === 'boolean') return value ? 1 : 0
  if (typeof value === 'string') {
    const cleaned = value.replace(/\s/g, '').replace(/,/g, '')
    if (!cleaned) return value
    const n = Number(cleaned)
    if (Number.isFinite(n)) return Math.round(n)
    return value
  }
  if (typeof value === 'object' && value !== null && 'result' in value) {
    return toIntegerValue((value as { result: CellValue }).result)
  }
  return value
}

/**
 * Organiza Contingente: omite vacías, detecta rótulos o solo datos,
 * recorta columnas, colorea y deja montos como enteros.
 */
export function writeOrganizedContingenteSheet(
  workbook: Workbook,
  source: Worksheet,
  sheetName = 'CONTINGENTE',
  variant: ContingenteVariant = 'standard',
): { sheet: Worksheet; dataRows: number } {
  const { keepSourceCols, colorByOutCol, integerOutCols } = layoutForVariant(variant)
  const integerCols = new Set(integerOutCols)

  const existing = workbook.getWorksheet(sheetName)
  if (existing) workbook.removeWorksheet(existing.id)

  const sheet = workbook.addWorksheet(sheetName)

  type PendingRow = { row: Row; isHeader: boolean }
  const nonEmpty: Row[] = []
  source.eachRow({ includeEmpty: true }, (row) => {
    if (isRowEmpty(row, keepSourceCols)) return
    nonEmpty.push(row)
  })

  let headerIdx = -1
  for (let i = 0; i < nonEmpty.length; i += 1) {
    if (isRotulosRow(nonEmpty[i], keepSourceCols)) {
      headerIdx = i
      break
    }
  }

  const pending: PendingRow[] = []
  if (headerIdx >= 0) {
    // Rótulos + todas las filas con datos debajo (títulos vacíos/anteriores se omiten)
    pending.push({ row: nonEmpty[headerIdx], isHeader: true })
    for (let i = headerIdx + 1; i < nonEmpty.length; i += 1) {
      pending.push({ row: nonEmpty[i], isHeader: false })
    }
  } else {
    // Sin rótulos: solo datos, sin inventar encabezado
    for (const row of nonEmpty) {
      pending.push({ row, isHeader: false })
    }
  }

  if (pending.length === 0) {
    throw new Error('La hoja Contingente no tiene filas útiles (solo vacías o sin datos).')
  }

  let outRow = 0
  let dataRows = 0
  let headerOutRow = 0

  for (const item of pending) {
    outRow += 1
    const newRow = sheet.getRow(outRow)
    if (item.isHeader) headerOutRow = outRow
    else dataRows += 1

    keepSourceCols.forEach((srcCol, idx) => {
      const outCol = idx + 1
      const srcCell = item.row.getCell(srcCol)
      const dest = newRow.getCell(outCol)
      const applyInteger = !item.isHeader && integerCols.has(outCol)
      if (applyInteger) {
        dest.value = toIntegerValue(srcCell.value)
        dest.numFmt = INTEGER_NUM_FMT
      } else {
        dest.value = srcCell.value
        if (srcCell.numFmt) dest.numFmt = srcCell.numFmt
      }
    })
    newRow.commit()
  }

  if (headerOutRow > 0) {
    const header = sheet.getRow(headerOutRow)
    header.font = { bold: true }
    header.commit()
  }

  for (let r = 1; r <= outRow; r += 1) {
    const isHeader = r === headerOutRow
    for (const [colStr, hex] of Object.entries(colorByOutCol)) {
      const col = Number(colStr)
      const cell = sheet.getRow(r).getCell(col)
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${hex}` },
      }
      if (!isHeader && integerCols.has(col) && cell.numFmt !== INTEGER_NUM_FMT) {
        cell.numFmt = INTEGER_NUM_FMT
      }
    }
  }

  keepSourceCols.forEach((_, idx) => {
    sheet.getColumn(idx + 1).width = idx === 5 || idx === 1 || idx === 6 ? 28 : 16
  })

  return { sheet, dataRows }
}

/** Todas las empresas usan Contingente organizado (con layout distinto en Cementos). */
export function empresaUsaContingenteOrganizado(_grupoCatalogo?: string): boolean {
  return true
}
