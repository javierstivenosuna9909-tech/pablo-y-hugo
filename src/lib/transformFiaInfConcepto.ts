import { mapHeaderIndexes, type RelacionRow } from './transformNomina'

/** Conceptos FIA (orden de columnas como en el ejemplo Inf x Concepto). */
export const FIA_CONCEPTOS = ['4CAA', '4CAF', '4CAR', '4L47', '4L63', '4CAQ'] as const

export type FiaConcepto = (typeof FIA_CONCEPTOS)[number]

export const FIA_INF_SHEET_NAME = 'Inf x Concepto'

export type FiaInfEmployeeRow = {
  nombre: string
  documento: string | number
  valores: Record<FiaConcepto, number | null>
  total: number
}

export type FiaInfResult = {
  rows: FiaInfEmployeeRow[]
  columnTotals: Record<FiaConcepto, number>
  grandTotal: number
}

export function empresaUsaFiaInfConcepto(empresaCodigo: string): boolean {
  return empresaCodigo.trim().toUpperCase() === 'CO05'
}

/** Acreedor FIA en catálogo Celsia. */
export const FIA_ACREEDOR = '103445'

export function isFiaAcreedor(acreedor: string): boolean {
  return acreedor.trim() === FIA_ACREEDOR
}

function cellToPrimitive(value: unknown): string | number {
  if (value == null) return ''
  if (typeof value === 'number' || typeof value === 'string') return value
  if (typeof value === 'boolean') return value ? 1 : 0
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object' && value !== null && 'result' in value) {
    return cellToPrimitive((value as { result: unknown }).result)
  }
  if (typeof value === 'object' && value !== null && 'text' in value) {
    return String((value as { text: unknown }).text ?? '')
  }
  return String(value)
}

function isBlank(value: unknown): boolean {
  if (value == null) return true
  if (typeof value === 'string') return value.trim() === ''
  return false
}

function toNeto(value: unknown): number {
  const raw = cellToPrimitive(value)
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw
  const n = Number(String(raw).replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : 0
}

function documentoKey(value: string | number): string {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Number.isInteger(value) ? String(value) : String(value)
  }
  return String(value).trim()
}

type FiaAcc = {
  nombre: string
  documento: string | number
  valores: Record<FiaConcepto, number>
}

function emptyColumnTotals(): Record<FiaConcepto, number> {
  return Object.fromEntries(FIA_CONCEPTOS.map((c) => [c, 0])) as Record<FiaConcepto, number>
}

function finalizeFiaInf(byDoc: Map<string, FiaAcc>): FiaInfResult {
  const rows: FiaInfEmployeeRow[] = [...byDoc.values()]
    .map((acc) => {
      const valores = Object.fromEntries(
        FIA_CONCEPTOS.map((c) => {
          const v = acc.valores[c]
          return [c, v === 0 ? null : v]
        }),
      ) as Record<FiaConcepto, number | null>
      const total = FIA_CONCEPTOS.reduce((sum, c) => sum + (acc.valores[c] || 0), 0)
      return {
        nombre: acc.nombre,
        documento: acc.documento,
        valores,
        total,
      }
    })
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }))

  const columnTotals = Object.fromEntries(
    FIA_CONCEPTOS.map((c) => [
      c,
      rows.reduce((sum, row) => sum + (row.valores[c] ?? 0), 0),
    ]),
  ) as Record<FiaConcepto, number>

  const grandTotal = rows.reduce((sum, row) => sum + row.total, 0)
  return { rows, columnTotals, grandTotal }
}

function addFiaValor(
  byDoc: Map<string, FiaAcc>,
  nombre: string,
  documento: string | number,
  concepto: FiaConcepto,
  valor: number,
): void {
  const key = documentoKey(documento)
  if (!key) return
  let acc = byDoc.get(key)
  if (!acc) {
    acc = {
      nombre,
      documento,
      valores: emptyColumnTotals(),
    }
    byDoc.set(key, acc)
  } else if (!acc.nombre && nombre) {
    acc.nombre = nombre
  }
  acc.valores[concepto] += valor
}

/**
 * Desde filas RELACION (conceptos FIA): pivot empleado × concepto + totales.
 */
export function buildFiaInfConceptoFromRelacionRows(rows: RelacionRow[]): FiaInfResult {
  const fiaSet = new Set<string>(FIA_CONCEPTOS)
  const byDoc = new Map<string, FiaAcc>()

  for (const row of rows) {
    const concepto = row.concepto.trim().toUpperCase()
    if (!fiaSet.has(concepto)) continue
    const nombre = row.nombreCompleto.trim()
    if (!nombre) continue
    addFiaValor(
      byDoc,
      nombre,
      row.cedula,
      concepto as FiaConcepto,
      toNeto(row.neto),
    )
  }

  if (byDoc.size === 0) {
    return { rows: [], columnTotals: emptyColumnTotals(), grandTotal: 0 }
  }
  return finalizeFiaInf(byDoc)
}

/**
 * Desde NOMINA cruda: pivot empleado × conceptos FIA + total por fila/columna.
 */
export function buildFiaInfConceptoFromNominaMatrix(matrix: unknown[][]): FiaInfResult {
  if (matrix.length === 0) {
    return { rows: [], columnTotals: emptyColumnTotals(), grandTotal: 0 }
  }

  let headerRowIndex = 0
  let indexes: ReturnType<typeof mapHeaderIndexes> | null = null
  for (let r = 0; r < Math.min(matrix.length, 20); r += 1) {
    try {
      indexes = mapHeaderIndexes(matrix[r] ?? [])
      headerRowIndex = r
      break
    } catch {
      // seguir
    }
  }
  if (!indexes) {
    throw new Error('No se encontraron columnas requeridas en NOMINA para Inf x Concepto.')
  }

  const fiaSet = new Set<string>(FIA_CONCEPTOS)
  const byDoc = new Map<string, FiaAcc>()

  for (let i = headerRowIndex + 1; i < matrix.length; i += 1) {
    const row = matrix[i] ?? []
    const nombre = String(cellToPrimitive(row[indexes.nombre]) ?? '').trim()
    if (isBlank(nombre)) continue

    const concepto = String(cellToPrimitive(row[indexes.concepto]) ?? '')
      .trim()
      .toUpperCase()
    if (!fiaSet.has(concepto)) continue

    const docRaw = cellToPrimitive(row[indexes.cedula])
    const doc = typeof docRaw === 'number' || typeof docRaw === 'string' ? docRaw : String(docRaw)
    addFiaValor(byDoc, nombre, doc, concepto as FiaConcepto, toNeto(row[indexes.neto]))
  }

  return finalizeFiaInf(byDoc)
}
