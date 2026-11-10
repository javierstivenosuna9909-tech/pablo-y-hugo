import { cellToDate } from './transformNomina'
import type { PagoTercerosRow } from './transformPagoTerceros'

export const CXP_SHEET_NAME = 'CXP'

export type CxpDocInfo = {
  numeroDocumento: string
  claseDocumento: string
  fechaDocumento: string
}

export type EnrichCxpResult = {
  rows: PagoTercerosRow[]
  updatedRows: number
  rowsWithDocumento: number
  cxpRowsUsed: number
  matchedByAcreedor: number
  filledWithCompanyDefault: number
  leftEmpty: number
  usedSociedadFallback: boolean
  missingHeaders: string[]
}

type CxpIndexes = {
  sociedad: number
  proveedor: number
  numeroDocumento: number
  claseDocumento: number
  fechaDocumento: number
}

const HEADER_ALIASES = {
  sociedad: ['sociedad'],
  // nit antes de tercero: layout contable Soluciones Modulares
  proveedor: ['nit', 'proveedor', 'tercero'],
  numeroDocumento: [
    'n documento',
    'no documento',
    'nro documento',
    'num documento',
    'numero documento',
    'nº documento',
    'n° documento',
    'nro de documento',
    'numero de documento',
    'consecutivo',
    'doc',
  ],
  claseDocumento: ['clase de documento', 'clase documento', 'tipo doc'],
  // Fe.contabilización primero: es la fuente oficial para Fecha documento / Fecha recibe CXP
  fechaDocumento: [
    'fe contabilizacion',
    'fecha contabilizacion',
    'fecha de contabilizacion',
    'fe.contabilizacion',
    'fecha de documento',
    'fecha documento',
    'fecha',
  ],
} as const

function headerCellText(value: unknown): string {
  if (value == null) return ''
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value)
  }
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object') {
    const obj = value as { text?: unknown; result?: unknown; richText?: Array<{ text?: string }> }
    if (typeof obj.text === 'string') return obj.text
    if (obj.result != null) return headerCellText(obj.result)
    if (Array.isArray(obj.richText)) return obj.richText.map((p) => p.text ?? '').join('')
  }
  return String(value)
}

function normalizeHeader(value: unknown): string {
  return headerCellText(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[º°]/g, '')
    .replace(/\./g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function findColumnIndex(
  headers: string[],
  aliases: readonly string[],
  options?: { preferStartsWith?: string[] },
): number {
  const normalized = headers.map(normalizeHeader)
  for (const alias of aliases) {
    const target = normalizeHeader(alias)
    const exact = normalized.findIndex((h) => h === target)
    if (exact >= 0) return exact
  }
  for (const alias of aliases) {
    const target = normalizeHeader(alias)
    const partial = normalized.findIndex((h) => {
      if (!h) return false
      // Evitar que "documento" matchee "clase de documento" / "fecha de documento"
      if (options?.preferStartsWith?.length) {
        const okPrefix = options.preferStartsWith.some((p) => h.startsWith(normalizeHeader(p)))
        if (!okPrefix) return false
      }
      return h.includes(target) || target.includes(h)
    })
    if (partial >= 0 && normalized[partial].length > 0) return partial
  }
  return -1
}

function cellText(value: unknown): string {
  if (value == null) return ''
  if (value instanceof Date) return formatFechaDocumento(value)
  if (typeof value === 'number' || typeof value === 'string') return String(value).trim()
  if (typeof value === 'boolean') return value ? '1' : '0'
  if (typeof value === 'object' && 'result' in (value as object)) {
    return cellText((value as { result: unknown }).result)
  }
  if (typeof value === 'object' && 'text' in (value as object)) {
    return String((value as { text: unknown }).text ?? '').trim()
  }
  if (typeof value === 'object' && 'richText' in (value as object)) {
    return headerCellText(value).trim()
  }
  return String(value).trim()
}

function formatFechaDocumento(value: unknown): string {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const dd = String(value.getUTCDate()).padStart(2, '0')
    const mm = String(value.getUTCMonth() + 1).padStart(2, '0')
    const yyyy = value.getUTCFullYear()
    return `${dd}/${mm}/${yyyy}`
  }
  const date = cellToDate(value)
  if (date) {
    const dd = String(date.getUTCDate()).padStart(2, '0')
    const mm = String(date.getUTCMonth() + 1).padStart(2, '0')
    const yyyy = date.getUTCFullYear()
    return `${dd}/${mm}/${yyyy}`
  }
  if (value == null || value === '') return ''
  if (typeof value === 'object' && 'result' in (value as object)) {
    return formatFechaDocumento((value as { result: unknown }).result)
  }
  if (typeof value === 'number' || typeof value === 'string') {
    const raw = String(value).trim()
    if (/invalid date/i.test(raw)) return ''
    return raw
  }
  return ''
}

function mostCommon(values: string[]): string {
  const counts = new Map<string, number>()
  for (const value of values) {
    if (!value) continue
    counts.set(value, (counts.get(value) ?? 0) + 1)
  }
  let best = ''
  let bestCount = 0
  for (const [value, count] of counts) {
    if (count > bestCount) {
      best = value
      bestCount = count
    }
  }
  return best
}

function readOptional(row: unknown[], index: number, asFecha = false): string {
  if (index < 0) return ''
  return asFecha ? formatFechaDocumento(row[index]) : cellText(row[index])
}

function collectFromRows(
  matrix: unknown[][],
  indexes: CxpIndexes,
  empresaFilter: string | null,
): {
  byProveedor: Map<string, CxpDocInfo>
  companyDefault: CxpDocInfo | null
  rowsUsed: number
} {
  const byProveedor = new Map<string, CxpDocInfo>()
  const companyDocs: string[] = []
  const companyClases: string[] = []
  const companyFechas: string[] = []
  let rowsUsed = 0
  const empresa = empresaFilter?.trim().toUpperCase() ?? null
  const canFilterSociedad = indexes.sociedad >= 0 && Boolean(empresa)

  for (let i = 1; i < matrix.length; i += 1) {
    const row = matrix[i] ?? []

    if (canFilterSociedad) {
      const sociedad = cellText(row[indexes.sociedad]).toUpperCase()
      if (sociedad && sociedad !== empresa) continue
    }

    const proveedor = readOptional(row, indexes.proveedor)
    const numeroDocumento = readOptional(row, indexes.numeroDocumento)
    const claseDocumento = readOptional(row, indexes.claseDocumento)
    const fechaDocumento = readOptional(row, indexes.fechaDocumento, true)

    if (!numeroDocumento && !claseDocumento && !fechaDocumento) continue

    rowsUsed += 1
    if (numeroDocumento) companyDocs.push(numeroDocumento)
    if (claseDocumento) companyClases.push(claseDocumento)
    if (fechaDocumento) companyFechas.push(fechaDocumento)

    // Solo matchear por acreedor si existe columna Proveedor
    if (indexes.proveedor < 0 || !proveedor) continue
    if (byProveedor.has(proveedor)) continue

    byProveedor.set(proveedor, {
      numeroDocumento,
      claseDocumento,
      fechaDocumento,
    })
  }

  const companyDefault: CxpDocInfo | null =
    rowsUsed === 0
      ? null
      : {
          numeroDocumento: mostCommon(companyDocs),
          claseDocumento: mostCommon(companyClases),
          fechaDocumento: mostCommon(companyFechas),
        }

  return { byProveedor, companyDefault, rowsUsed }
}

export function extractCxpDocs(
  matrix: unknown[][],
  empresaCodigo: string,
): {
  byProveedor: Map<string, CxpDocInfo>
  companyDefault: CxpDocInfo | null
  rowsUsed: number
  usedSociedadFallback: boolean
  missingHeaders: string[]
  indexes: CxpIndexes
} {
  if (matrix.length === 0) {
    throw new Error('La pestaña CXP está vacía.')
  }

  const headers = (matrix[0] ?? []).map((h) => headerCellText(h))
  const indexes: CxpIndexes = {
    sociedad: findColumnIndex(headers, HEADER_ALIASES.sociedad),
    proveedor: findColumnIndex(headers, HEADER_ALIASES.proveedor),
    numeroDocumento: findColumnIndex(headers, HEADER_ALIASES.numeroDocumento, {
      preferStartsWith: ['n ', 'no ', 'nro', 'num', 'numero'],
    }),
    claseDocumento: findColumnIndex(headers, HEADER_ALIASES.claseDocumento, {
      preferStartsWith: ['clase', 'tipo'],
    }),
    fechaDocumento: findColumnIndex(headers, HEADER_ALIASES.fechaDocumento, {
      preferStartsWith: ['fe contabil', 'fecha contabil', 'fecha de', 'fecha', 'fe'],
    }),
  }

  const missingHeaders: string[] = []
  if (indexes.numeroDocumento < 0) missingHeaders.push('Nº documento')
  if (indexes.claseDocumento < 0) missingHeaders.push('Clase de documento')
  if (indexes.fechaDocumento < 0) missingHeaders.push('Fecha de documento')

  const hasDocColumn =
    indexes.numeroDocumento >= 0 ||
    indexes.claseDocumento >= 0 ||
    indexes.fechaDocumento >= 0

  // Homologación solo-Proveedor (Casa para mí): aceptar sin inventar documento/clase/fecha
  if (!hasDocColumn) {
    if (indexes.proveedor < 0) {
      throw new Error(
        `CXP no trae columnas útiles (Nº documento / Clase / Fecha / Proveedor-NIT). Encabezados: ${headers.filter(Boolean).join(' | ')}`,
      )
    }
    let rowsUsed = 0
    for (let i = 1; i < matrix.length; i += 1) {
      const proveedor = readOptional(matrix[i] ?? [], indexes.proveedor)
      if (proveedor) rowsUsed += 1
    }
    return {
      byProveedor: new Map(),
      companyDefault: {
        numeroDocumento: '',
        claseDocumento: '',
        fechaDocumento: '',
      },
      rowsUsed,
      usedSociedadFallback: false,
      missingHeaders,
      indexes,
    }
  }

  // 1) Filtrar por sociedad solo si existe esa columna
  let result = collectFromRows(matrix, indexes, indexes.sociedad >= 0 ? empresaCodigo : null)
  let usedSociedadFallback = false

  // 2) Si filtró por sociedad y no hubo filas, usar todo el CXP
  if (result.rowsUsed === 0 && indexes.sociedad >= 0) {
    result = collectFromRows(matrix, indexes, null)
    usedSociedadFallback = result.rowsUsed > 0
  }

  return { ...result, usedSociedadFallback, missingHeaders, indexes }
}

export function enrichPagoFromCxp(
  pagoRows: PagoTercerosRow[],
  cxpMatrix: unknown[][],
  empresaCodigo: string,
): EnrichCxpResult {
  const { byProveedor, companyDefault, rowsUsed, usedSociedadFallback, missingHeaders, indexes } =
    extractCxpDocs(cxpMatrix, empresaCodigo)

  const hasDocColumn =
    indexes.numeroDocumento >= 0 ||
    indexes.claseDocumento >= 0 ||
    indexes.fechaDocumento >= 0

  // Solo-Proveedor: carga válida, sin completar documento/clase/fecha
  if (!hasDocColumn) {
    return {
      rows: pagoRows,
      updatedRows: 0,
      rowsWithDocumento: 0,
      cxpRowsUsed: rowsUsed,
      matchedByAcreedor: 0,
      filledWithCompanyDefault: 0,
      leftEmpty: pagoRows.length,
      usedSociedadFallback,
      missingHeaders,
    }
  }

  if (rowsUsed === 0 || !companyDefault) {
    throw new Error(
      `No hay filas útiles en CXP para completar documento/clase/fecha (empresa ${empresaCodigo}).`,
    )
  }

  let matchedByAcreedor = 0
  let filledWithCompanyDefault = 0
  let leftEmpty = 0
  let updatedRows = 0
  let rowsWithDocumento = 0

  const rows = pagoRows.map((row) => {
    const acreedor = row.acreedor.trim()
    const canMatchProveedor = indexes.proveedor >= 0
    const fromProveedor =
      canMatchProveedor && acreedor && acreedor.toUpperCase() !== 'N/A'
        ? byProveedor.get(acreedor)
        : undefined

    let info: CxpDocInfo
    if (fromProveedor) {
      info = {
        numeroDocumento: fromProveedor.numeroDocumento || companyDefault.numeroDocumento,
        claseDocumento: fromProveedor.claseDocumento || companyDefault.claseDocumento,
        fechaDocumento: fromProveedor.fechaDocumento || companyDefault.fechaDocumento,
      }
      matchedByAcreedor += 1
    } else {
      info = { ...companyDefault }
      filledWithCompanyDefault += 1
    }

    // Solo llenar campos cuyo encabezado existía; el resto queda en blanco
    const numeroDocumento =
      indexes.numeroDocumento >= 0 ? info.numeroDocumento || row.numeroDocumento || '' : ''
    const claseDocumento =
      indexes.claseDocumento >= 0 ? info.claseDocumento || row.claseDocumento || '' : ''
    const fechaDocumento =
      indexes.fechaDocumento >= 0 ? info.fechaDocumento || row.fechaDocumento || '' : ''
    const fechaRecibeCxp = fechaDocumento

    const next: PagoTercerosRow = {
      ...row,
      numeroDocumento,
      claseDocumento,
      fechaDocumento,
      fechaRecibeCxp,
    }

    if (next.numeroDocumento) rowsWithDocumento += 1
    if (!next.numeroDocumento && !next.claseDocumento && !next.fechaDocumento) leftEmpty += 1

    if (
      next.numeroDocumento !== row.numeroDocumento ||
      next.claseDocumento !== row.claseDocumento ||
      next.fechaDocumento !== row.fechaDocumento ||
      next.fechaRecibeCxp !== row.fechaRecibeCxp
    ) {
      updatedRows += 1
    }

    return next
  })

  return {
    rows,
    updatedRows,
    rowsWithDocumento,
    cxpRowsUsed: rowsUsed,
    matchedByAcreedor,
    filledWithCompanyDefault,
    leftEmpty,
    usedSociedadFallback,
    missingHeaders,
  }
}
