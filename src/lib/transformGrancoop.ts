import type { RelacionRow } from './transformNomina'

/** Acreedor GRANCOOP / FRV en catálogo Celsia. */
export const GRANCOOP_ACREEDOR = '104530'

/** Conceptos GRANCOOP en orden de columnas TD X CONCEPTO (ejemplo). */
export const GRANCOOP_TD_CONCEPTOS = [
  '4CBH',
  '4CAG',
  '4CAS',
  '4CAB',
  '4L37',
  '4CBI',
  '4CDG',
  '4CDS',
] as const

export type GrancoopTdConcepto = (typeof GRANCOOP_TD_CONCEPTOS)[number]

/** Conceptos del Centralizador que alimentan CARTERA / INFORME FINAL. */
export const GRANCOOP_CARTERA_CONCEPTOS = ['4CBH', '4CBI'] as const

export function empresaUsaGrancoopCentralizador(empresaCodigo: string): boolean {
  return empresaCodigo.trim().toUpperCase() === 'CO05'
}

export function isGrancoopAcreedor(acreedor: string): boolean {
  return acreedor.trim() === GRANCOOP_ACREEDOR
}

export type CentralizadorNovedad = {
  documento: string | number
  nombre: string
  concepto: string
  monto: number
  compania: string
}

export type GrancoopCarteraRow = {
  cedula: string | number
  nombre: string
  concepto: string
  descuento: number
}

export type GrancoopNominaRow = {
  empleado: string | number
  cedula: string | number
  nombreCompleto: string
  concepto: string
  descripcion: string
  descuentos: number
  sociedad: string
}

export type GrancoopTdRow = {
  cedula: string | number
  nombreCompleto: string
  valores: Record<GrancoopTdConcepto, number | null>
  total: number
}

export type GrancoopInformeRow = {
  empleado: string | number
  colaborador: string
  concepto: string
  sumaValor: number
  valorEnNomina: number
  diferencia: number
  observaciones: string
}

export type GrancoopBundle = {
  nomina: GrancoopNominaRow[]
  td: GrancoopTdRow[]
  tdColumnTotals: Record<GrancoopTdConcepto, number>
  tdGrandTotal: number
  cartera: GrancoopCarteraRow[]
  informe: GrancoopInformeRow[]
}

function normalizeText(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/\s+/g, ' ')
    .trim()
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
  if (typeof value === 'object' && value !== null && 'richText' in value) {
    const parts = (value as { richText: Array<{ text?: string }> }).richText
    return parts.map((p) => p.text ?? '').join('')
  }
  return String(value)
}

function documentoKey(value: string | number): string {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Number.isInteger(value) ? String(value) : String(value)
  }
  return String(value).trim()
}

function toNeto(value: unknown): number {
  const raw = cellToPrimitive(value)
  if (typeof raw === 'number' && Number.isFinite(raw)) return raw
  const n = Number(String(raw).replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : 0
}

/** `COP 25420.00` / `25420` → número positivo. */
export function parseCentralizadorMonto(value: unknown): number {
  const raw = String(cellToPrimitive(value) ?? '').trim()
  if (!raw) return 0
  const cleaned = raw
    .replace(/^COP\s*/i, '')
    .replace(/\s/g, '')
    .replace(/,/g, '')
  const n = Number(cleaned)
  return Number.isFinite(n) ? n : 0
}

/** `4CBH - CUOTA...` → `4CBH`. */
export function parseCentralizadorConcepto(value: unknown): string {
  const raw = String(cellToPrimitive(value) ?? '').trim()
  if (!raw) return ''
  const code = raw.split('-')[0]?.trim() ?? raw
  return code.toUpperCase()
}

export function isCelsiaColombiaCompany(compania: string): boolean {
  return normalizeText(compania).includes('CELSIA COLOMBIA')
}

function findHeaderRow(matrix: unknown[][]): { rowIndex: number; headers: string[] } | null {
  for (let r = 0; r < Math.min(matrix.length, 30); r += 1) {
    const row = matrix[r] ?? []
    const headers = row.map((c) => normalizeText(String(cellToPrimitive(c) ?? '')))
    const hasId = headers.some((h) => h === 'ID')
    const hasDoc = headers.some((h) => h.includes('DOCUMENTO'))
    const hasConcepto = headers.some((h) => h.includes('CONCEPTO'))
    const hasMonto = headers.some((h) => h.includes('MONTO') || h.includes('CANTIDAD'))
    if (hasId && hasDoc && hasConcepto && hasMonto) {
      return { rowIndex: r, headers }
    }
  }
  return null
}

function colIndex(headers: string[], ...aliases: string[]): number {
  for (const alias of aliases) {
    const target = normalizeText(alias)
    const exact = headers.findIndex((h) => h === target)
    if (exact >= 0) return exact
  }
  for (const alias of aliases) {
    const target = normalizeText(alias)
    const partial = headers.findIndex((h) => h.includes(target))
    if (partial >= 0) return partial
  }
  return -1
}

/**
 * Lee pestaña Novedades del Centralizador (filas de datos tras encabezados).
 * No filtra aún por compañía/concepto.
 */
export function parseCentralizadorNovedades(matrix: unknown[][]): CentralizadorNovedad[] {
  const found = findHeaderRow(matrix)
  if (!found) {
    throw new Error(
      'No se encontró la fila de encabezados en Centralizador (ID, Documento, Concepto, Monto/Cantidad).',
    )
  }
  const { rowIndex, headers } = found
  const iCompania = colIndex(headers, 'COMPANIA', 'COMPAÑIA')
  const iDoc = colIndex(headers, 'DOCUMENTO')
  const iNombre = colIndex(headers, 'NOMBRE DEL COLABORADOR', 'NOMBRE')
  const iConcepto = colIndex(headers, 'CONCEPTO')
  const iMonto = colIndex(headers, 'MONTO/CANTIDAD', 'MONTO', 'CANTIDAD')

  if (iDoc < 0 || iConcepto < 0 || iMonto < 0) {
    throw new Error(
      'Centralizador incompleto: faltan columnas Documento, Concepto o Monto/Cantidad.',
    )
  }

  const out: CentralizadorNovedad[] = []
  for (let r = rowIndex + 1; r < matrix.length; r += 1) {
    const row = matrix[r] ?? []
    const concepto = parseCentralizadorConcepto(row[iConcepto])
    if (!concepto) continue
    const docRaw = cellToPrimitive(row[iDoc])
    if (docRaw === '' || docRaw == null) continue
    const monto = parseCentralizadorMonto(row[iMonto])
    out.push({
      documento: docRaw,
      nombre: String(cellToPrimitive(iNombre >= 0 ? row[iNombre] : '') ?? '').trim(),
      concepto,
      monto,
      compania: String(cellToPrimitive(iCompania >= 0 ? row[iCompania] : '') ?? '').trim(),
    })
  }
  return out
}

/** Filas Centralizador → CARTERA (CO05 / 4CBH+4CBI), agregadas por cédula+concepto. */
export function buildGrancoopCartera(
  novedades: CentralizadorNovedad[],
): GrancoopCarteraRow[] {
  const allowed = new Set<string>(GRANCOOP_CARTERA_CONCEPTOS)
  const byKey = new Map<string, GrancoopCarteraRow>()

  for (const n of novedades) {
    if (!isCelsiaColombiaCompany(n.compania)) continue
    if (!allowed.has(n.concepto)) continue
    const key = `${documentoKey(n.documento)}|${n.concepto}`
    const current = byKey.get(key)
    if (current) {
      current.descuento += n.monto
      if (!current.nombre && n.nombre) current.nombre = n.nombre
    } else {
      byKey.set(key, {
        cedula: n.documento,
        nombre: n.nombre,
        concepto: n.concepto,
        descuento: n.monto,
      })
    }
  }

  return [...byKey.values()].sort((a, b) => {
    const byDoc = documentoKey(a.cedula).localeCompare(documentoKey(b.cedula), 'es', {
      numeric: true,
    })
    if (byDoc !== 0) return byDoc
    return a.concepto.localeCompare(b.concepto, 'es')
  })
}

export function buildGrancoopNominaRows(
  relacionRows: RelacionRow[],
  sociedadFallback: string,
): GrancoopNominaRow[] {
  const allowed = new Set<string>(GRANCOOP_TD_CONCEPTOS)
  const rows: GrancoopNominaRow[] = []
  for (const row of relacionRows) {
    const concepto = row.concepto.trim().toUpperCase()
    if (!allowed.has(concepto)) continue
    rows.push({
      empleado: row.empleado ?? '',
      cedula: row.cedula,
      nombreCompleto: row.nombreCompleto,
      concepto,
      descripcion: row.descripcion,
      descuentos: toNeto(row.neto),
      sociedad: row.sociedad || sociedadFallback,
    })
  }
  return rows.sort((a, b) => {
    const byEmp = String(a.empleado).localeCompare(String(b.empleado), 'es', { numeric: true })
    if (byEmp !== 0) return byEmp
    const byDoc = documentoKey(a.cedula).localeCompare(documentoKey(b.cedula), 'es', {
      numeric: true,
    })
    if (byDoc !== 0) return byDoc
    return a.concepto.localeCompare(b.concepto, 'es')
  })
}

export function buildGrancoopTdFromNomina(nomina: GrancoopNominaRow[]): {
  rows: GrancoopTdRow[]
  columnTotals: Record<GrancoopTdConcepto, number>
  grandTotal: number
} {
  type Acc = {
    cedula: string | number
    nombreCompleto: string
    valores: Record<GrancoopTdConcepto, number>
  }
  const byDoc = new Map<string, Acc>()

  for (const row of nomina) {
    const concepto = row.concepto.trim().toUpperCase() as GrancoopTdConcepto
    if (!GRANCOOP_TD_CONCEPTOS.includes(concepto)) continue
    const key = documentoKey(row.cedula)
    if (!key) continue
    let acc = byDoc.get(key)
    if (!acc) {
      acc = {
        cedula: row.cedula,
        nombreCompleto: row.nombreCompleto,
        valores: Object.fromEntries(GRANCOOP_TD_CONCEPTOS.map((c) => [c, 0])) as Record<
          GrancoopTdConcepto,
          number
        >,
      }
      byDoc.set(key, acc)
    } else if (!acc.nombreCompleto && row.nombreCompleto) {
      acc.nombreCompleto = row.nombreCompleto
    }
    acc.valores[concepto] += row.descuentos
  }

  const rows: GrancoopTdRow[] = [...byDoc.values()]
    .map((acc) => {
      const valores = Object.fromEntries(
        GRANCOOP_TD_CONCEPTOS.map((c) => {
          const v = acc.valores[c]
          return [c, v === 0 ? null : v]
        }),
      ) as Record<GrancoopTdConcepto, number | null>
      const total = GRANCOOP_TD_CONCEPTOS.reduce((sum, c) => sum + (acc.valores[c] || 0), 0)
      return {
        cedula: acc.cedula,
        nombreCompleto: acc.nombreCompleto,
        valores,
        total,
      }
    })
    .sort((a, b) =>
      documentoKey(a.cedula).localeCompare(documentoKey(b.cedula), 'es', { numeric: true }),
    )

  const columnTotals = Object.fromEntries(
    GRANCOOP_TD_CONCEPTOS.map((c) => [
      c,
      rows.reduce((sum, row) => sum + (row.valores[c] ?? 0), 0),
    ]),
  ) as Record<GrancoopTdConcepto, number>

  const grandTotal = rows.reduce((sum, row) => sum + row.total, 0)
  return { rows, columnTotals, grandTotal }
}

function observacionesInforme(sumaCentralizador: number, valorNomina: number): string {
  const diff = valorNomina + sumaCentralizador
  if (diff === 0) return ''
  const absNom = Math.abs(valorNomina)
  const absCen = Math.abs(sumaCentralizador)
  if (absNom < absCen) return 'TIENE PROVISION DE VACACIONES'
  if (absNom > absCen) return 'PROVISION DE VACACIONES'
  return ''
}

export function buildGrancoopInforme(
  cartera: GrancoopCarteraRow[],
  nomina: GrancoopNominaRow[],
): GrancoopInformeRow[] {
  const nominaByKey = new Map<string, { valor: number; empleado: string | number; nombre: string }>()
  for (const row of nomina) {
    const key = `${documentoKey(row.cedula)}|${row.concepto.trim().toUpperCase()}`
    const current = nominaByKey.get(key)
    if (current) {
      current.valor += row.descuentos
      if (!current.empleado && row.empleado) current.empleado = row.empleado
    } else {
      nominaByKey.set(key, {
        valor: row.descuentos,
        empleado: row.empleado || row.cedula,
        nombre: row.nombreCompleto,
      })
    }
  }

  return cartera.map((c) => {
    const key = `${documentoKey(c.cedula)}|${c.concepto.trim().toUpperCase()}`
    const nom = nominaByKey.get(key)
    const valorEnNomina = nom?.valor ?? 0
    const sumaValor = c.descuento
    const diferencia = valorEnNomina + sumaValor
    return {
      empleado: nom?.empleado || c.cedula,
      colaborador: (nom?.nombre || c.nombre).trim(),
      concepto: c.concepto,
      sumaValor,
      valorEnNomina,
      diferencia,
      observaciones: observacionesInforme(sumaValor, valorEnNomina),
    }
  })
}

/** Arma el paquete completo GRANCOOP para el Excel del ZIP. */
export function buildGrancoopBundle(
  relacionRows: RelacionRow[],
  novedades: CentralizadorNovedad[],
  sociedadFallback: string,
): GrancoopBundle {
  const nomina = buildGrancoopNominaRows(relacionRows, sociedadFallback)
  const td = buildGrancoopTdFromNomina(nomina)
  const cartera = buildGrancoopCartera(novedades)
  const informe = buildGrancoopInforme(cartera, nomina)
  return {
    nomina,
    td: td.rows,
    tdColumnTotals: td.columnTotals,
    tdGrandTotal: td.grandTotal,
    cartera,
    informe,
  }
}
