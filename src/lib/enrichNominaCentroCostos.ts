import type { Cell, CellValue, Workbook, Worksheet } from 'exceljs'

export const MASTER_BD_SHEET_NAME = 'BD COMPLETA'
export const CENTRO_COSTOS_HEADER = 'CENTRO DE COSTOS'

/** Columna I (1-based) */
const INSERT_COL = 9
/** Columna A: Número de personal */
const PERSONAL_COL = 1

function normalizeHeader(value: unknown): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ')
}

export function cellToKey(value: CellValue | unknown): string {
  if (value == null || value === '') return ''
  if (typeof value === 'number' && Number.isFinite(value)) {
    // Evitar "123.0" vs "123"
    return Number.isInteger(value) ? String(value) : String(value)
  }
  if (typeof value === 'boolean') return value ? '1' : '0'
  if (typeof value === 'string') return value.trim()
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object' && value !== null && 'result' in value) {
    return cellToKey((value as { result: CellValue }).result)
  }
  if (typeof value === 'object' && value !== null && 'text' in value) {
    return String((value as { text: unknown }).text ?? '').trim()
  }
  return String(value).trim()
}

/** INO/ICP → FUNDACION; otro centro de costos no vacío → GANADERIA. */
export function mapCostCenterToCecoLabel(costCenter: unknown): string {
  const raw = cellToKey(costCenter)
  if (!raw) return ''
  const upper = raw.toUpperCase()
  if (upper.includes('INO') || upper.includes('ICP')) return 'FUNDACION'
  return 'GANADERIA'
}

/** Mapa Número de personal / ASSIGNMENT ID → CECO (FUNDACION|GANADERIA). */
export function buildCecoLabelByPersonal(
  costByAssignment: Map<string, CellValue>,
): Map<string, string> {
  const out = new Map<string, string>()
  for (const [personal, cost] of costByAssignment) {
    const label = mapCostCenterToCecoLabel(cost)
    if (label) out.set(personal, label)
  }
  return out
}

function copyCell(from: Cell, to: Cell): void {
  to.value = from.value
  if (from.style) to.style = { ...from.style }
  if (from.numFmt) to.numFmt = from.numFmt
}

function sheetMaxColumn(sheet: Worksheet): number {
  let max = 0
  sheet.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (_cell, col) => {
      if (col > max) max = col
    })
  })
  return max
}

/** Mapa ASSIGNMENT ID → COST CENTER CODE desde BD COMPLETA. */
export function buildAssignmentCostCenterMap(master: Workbook): Map<string, CellValue> {
  const bd =
    master.worksheets.find(
      (ws) => normalizeHeader(ws.name) === normalizeHeader(MASTER_BD_SHEET_NAME),
    ) ?? null
  if (!bd) {
    const names = master.worksheets.map((ws) => ws.name).join(', ')
    throw new Error(
      `No se encontró la pestaña "${MASTER_BD_SHEET_NAME}" en MASTER. Hojas: ${names || '(ninguna)'}`,
    )
  }

  // ASSIGNMENT ID = col B, COST CENTER CODE = col S; datos desde fila 4
  const assignmentCol = 2
  const costCol = 19
  const dataStartRow = 4

  const map = new Map<string, CellValue>()
  bd.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber < dataStartRow) return
    const key = cellToKey(row.getCell(assignmentCol).value)
    if (!key || map.has(key)) return
    map.set(key, row.getCell(costCol).value)
  })
  return map
}

/** Fila de encabezados reales (p. ej. fila 3 si hay título arriba). */
function findNominaHeaderRow(sheet: Worksheet): number {
  const maxScan = Math.min(20, sheet.rowCount || 20)
  for (let r = 1; r <= maxScan; r += 1) {
    const row = sheet.getRow(r)
    let found = false
    row.eachCell({ includeEmpty: false }, (cell) => {
      const n = normalizeHeader(cell.value)
      if (
        n.includes('numero de personal') ||
        n.includes('nombre empl') ||
        n.includes('denominacion') ||
        n === 'subtipo' ||
        n.includes('cc-nomina')
      ) {
        found = true
      }
    })
    if (found) return r
  }
  return 1
}

/**
 * Inserta columna I "CENTRO DE COSTOS" en NOMINA y rellena desde el mapa
 * (Número de personal col A → COST CENTER CODE).
 */
export function enrichNominaWithCentroCostos(
  nominaSheet: Worksheet,
  costByAssignment: Map<string, CellValue>,
): { matched: number; totalDataRows: number } {
  const maxCol = Math.max(sheetMaxColumn(nominaSheet), INSERT_COL)
  const headerRowNum = findNominaHeaderRow(nominaSheet)

  // Desplazar columnas I..max hacia la derecha
  nominaSheet.eachRow({ includeEmpty: true }, (row) => {
    for (let c = maxCol; c >= INSERT_COL; c -= 1) {
      const src = row.getCell(c)
      const dest = row.getCell(c + 1)
      copyCell(src, dest)
    }
    const inserted = row.getCell(INSERT_COL)
    inserted.value = null
    inserted.style = {}
  })

  // Encabezado en la misma fila que el resto de títulos (no en la fila 1 de título)
  const headerRow = nominaSheet.getRow(headerRowNum)
  const headerCell = headerRow.getCell(INSERT_COL)
  headerCell.value = CENTRO_COSTOS_HEADER
  const styleRef = headerRow.getCell(INSERT_COL - 1)
  if (styleRef.style && Object.keys(styleRef.style).length > 0) {
    headerCell.style = { ...styleRef.style }
  } else {
    headerCell.font = { bold: true }
  }

  let matched = 0
  let totalDataRows = 0

  nominaSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (rowNumber <= headerRowNum) return
    const personal = cellToKey(row.getCell(PERSONAL_COL).value)
    if (!personal) {
      row.getCell(INSERT_COL).value = ''
      return
    }
    // Evitar tratar textos de cabecera/título como número de personal
    if (/[a-zA-ZáéíóúÁÉÍÓÚ]/.test(personal) && !/^\d+$/.test(personal)) {
      return
    }
    totalDataRows += 1
    if (costByAssignment.has(personal)) {
      row.getCell(INSERT_COL).value = costByAssignment.get(personal) ?? ''
      matched += 1
    } else {
      row.getCell(INSERT_COL).value = ''
    }
  })

  return { matched, totalDataRows }
}

export function empresaUsaCentroCostosMaster(empresaCodigo: string): boolean {
  return empresaCodigo.trim() === '1600'
}
