export const RELACION_HEADERS = [
  'AÑO',
  'MES',
  'PERIODO',
  'CEDULA',
  'NOMBRE_COMPLETO',
  'CONCEPTO',
  'DESCRIPCION',
  'NETO',
  'SOCIEDAD',
] as const

export type RelacionRow = {
  anio: number
  mes: number
  periodo: string | number
  cedula: string | number
  nombreCompleto: string
  concepto: string
  descripcion: string
  neto: number | string
  sociedad: string
  /** Fundación Grupo Argos: FUNDACION | GANADERIA */
  ceco?: string
  /** Número de personal (col A NOMINA), p. ej. hoja GRANCOOP. */
  empleado?: string | number
}

export type TransformResult = {
  rows: RelacionRow[]
  detailRows: number
  skippedTotalRows: number
}

const REQUIRED_HEADERS = {
  nombre: ['nombre empl./cand.', 'nombre empl/cand', 'nombre'],
  cedula: ['número id', 'numero id', 'número de id', 'cedula', 'cédula'],
  concepto: ['cc-nómina', 'cc-nomina', 'cc nomina'],
  descripcion: ['texto expl.cc-nómina', 'texto expl.cc-nomina', 'texto expl cc-nomina'],
  neto: ['total'],
  fechaFin: ['per.para, fecha fin', 'per.para fecha fin', 'per para fecha fin'],
} as const

function normalizeHeader(value: unknown): string {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ')
}

function findColumnIndex(headers: string[], aliases: readonly string[]): number {
  const normalized = headers.map(normalizeHeader)
  for (const alias of aliases) {
    const target = normalizeHeader(alias)
    const exact = normalized.findIndex((h) => h === target)
    if (exact >= 0) return exact
  }
  for (const alias of aliases) {
    const target = normalizeHeader(alias)
    const partial = normalized.findIndex((h) => h.includes(target) || target.includes(h))
    if (partial >= 0 && normalized[partial].length > 0) return partial
  }
  return -1
}

/** Excel serial date → JS Date (1900 date system). */
export function excelSerialToDate(serial: number): Date {
  const utcDays = Math.floor(serial - 25569)
  const utcValue = utcDays * 86400
  return new Date(utcValue * 1000)
}

export function cellToDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    return excelSerialToDate(value)
  }
  if (typeof value === 'string' && value.trim() !== '') {
    const asNumber = Number(value)
    if (Number.isFinite(asNumber) && asNumber > 20000) {
      return excelSerialToDate(asNumber)
    }
    const parsed = new Date(value)
    if (!Number.isNaN(parsed.getTime())) return parsed
  }
  return null
}

export function periodoFromDate(date: Date): number {
  return date.getUTCDate() <= 15 ? 10 : 20
}

function cellToPrimitive(value: unknown): string | number {
  if (value == null) return ''
  if (typeof value === 'number' || typeof value === 'string') return value
  if (typeof value === 'boolean') return value ? 1 : 0
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object' && 'result' in (value as object)) {
    return cellToPrimitive((value as { result: unknown }).result)
  }
  if (typeof value === 'object' && 'text' in (value as object)) {
    return String((value as { text: unknown }).text ?? '')
  }
  return String(value)
}

function isBlank(value: unknown): boolean {
  if (value == null) return true
  if (typeof value === 'string') return value.trim() === ''
  return false
}

export function mapHeaderIndexes(headerRow: unknown[]): Record<keyof typeof REQUIRED_HEADERS, number> {
  const headers = headerRow.map((h) => String(h ?? ''))
  const indexes = {
    nombre: findColumnIndex(headers, REQUIRED_HEADERS.nombre),
    cedula: findColumnIndex(headers, REQUIRED_HEADERS.cedula),
    concepto: findColumnIndex(headers, REQUIRED_HEADERS.concepto),
    descripcion: findColumnIndex(headers, REQUIRED_HEADERS.descripcion),
    neto: findColumnIndex(headers, REQUIRED_HEADERS.neto),
    fechaFin: findColumnIndex(headers, REQUIRED_HEADERS.fechaFin),
  }

  const missing = Object.entries(indexes)
    .filter(([, idx]) => idx < 0)
    .map(([key]) => key)

  if (missing.length > 0) {
    throw new Error(
      `No se encontraron columnas requeridas en NOMINA: ${missing.join(', ')}. Encabezados: ${headers.filter(Boolean).join(' | ')}`,
    )
  }

  return indexes
}

/** Primera fecha fin válida → periodo (10/20). Null si no hay filas útiles. */
export function detectPeriodoFromNominaMatrix(matrix: unknown[][]): number | null {
  if (matrix.length === 0) return null
  let indexes: ReturnType<typeof mapHeaderIndexes>
  try {
    indexes = mapHeaderIndexes(matrix[0] ?? [])
  } catch {
    return null
  }

  for (let i = 1; i < matrix.length; i += 1) {
    const row = matrix[i] ?? []
    const nombre = cellToPrimitive(row[indexes.nombre])
    if (isBlank(nombre)) continue
    const fecha = cellToDate(row[indexes.fechaFin])
    if (!fecha) continue
    return periodoFromDate(fecha)
  }
  return null
}

function personalKeyFromCell(value: unknown): string {
  if (value == null || value === '') return ''
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Number.isInteger(value) ? String(value) : String(value)
  }
  if (typeof value === 'string') return value.trim()
  if (typeof value === 'object' && value !== null && 'result' in value) {
    return personalKeyFromCell((value as { result: unknown }).result)
  }
  if (typeof value === 'object' && value !== null && 'text' in value) {
    return String((value as { text: unknown }).text ?? '').trim()
  }
  return String(value).trim()
}

export function transformNominaRows(
  matrix: unknown[][],
  sociedad: string,
  periodoOverride?: string | number,
  /** Fundación 1600: personal (col A) → FUNDACION|GANADERIA */
  cecoByPersonal?: Map<string, string>,
): TransformResult {
  if (matrix.length === 0) {
    throw new Error('La pestaña NOMINA está vacía.')
  }

  // Encabezados: primera fila que tenga las columnas requeridas (puede no ser la 1)
  let headerRowIndex = 0
  let indexes: ReturnType<typeof mapHeaderIndexes> | null = null
  for (let r = 0; r < Math.min(matrix.length, 20); r += 1) {
    try {
      indexes = mapHeaderIndexes(matrix[r] ?? [])
      headerRowIndex = r
      break
    } catch {
      // seguir buscando
    }
  }
  if (!indexes) {
    mapHeaderIndexes(matrix[0] ?? [])
    throw new Error('No se encontraron columnas requeridas en NOMINA.')
  }

  const rows: RelacionRow[] = []
  let skippedTotalRows = 0
  const overrideText =
    periodoOverride == null ? '' : String(periodoOverride).trim()
  const hasOverride = overrideText.length > 0
  const overrideValue: string | number = (() => {
    if (!hasOverride) return ''
    const asNum = Number(overrideText)
    if (Number.isFinite(asNum) && String(asNum) === overrideText) return asNum
    return overrideText
  })()

  for (let i = headerRowIndex + 1; i < matrix.length; i += 1) {
    const row = matrix[i] ?? []
    const nombre = cellToPrimitive(row[indexes.nombre])

    if (isBlank(nombre)) {
      // Filas de total / subtotal sin nombre de empleado
      const hasAnyValue = row.some((cell) => !isBlank(cellToPrimitive(cell)))
      if (hasAnyValue) skippedTotalRows += 1
      continue
    }

    const fecha = cellToDate(row[indexes.fechaFin])
    if (!fecha) {
      throw new Error(
        `Fila ${i + 1}: no se pudo interpretar "Per.Para, fecha fin" (${String(row[indexes.fechaFin] ?? '')}).`,
      )
    }

    const netoRaw = cellToPrimitive(row[indexes.neto])
    const neto =
      typeof netoRaw === 'number'
        ? netoRaw
        : netoRaw === ''
          ? 0
          : Number.isFinite(Number(netoRaw))
            ? Number(netoRaw)
            : netoRaw

    // No incluir en RELACION (ni ZIP / PAGO) filas con NETO exactamente 0
    if (typeof neto === 'number' && neto === 0) continue

    const personal = personalKeyFromCell(row[0])
    const ceco =
      cecoByPersonal && personal ? cecoByPersonal.get(personal) : undefined
    const empleadoRaw = cellToPrimitive(row[0])

    rows.push({
      anio: fecha.getUTCFullYear(),
      mes: fecha.getUTCMonth() + 1,
      periodo: hasOverride ? overrideValue : periodoFromDate(fecha),
      cedula: cellToPrimitive(row[indexes.cedula]),
      nombreCompleto: String(nombre).trim(),
      concepto: String(cellToPrimitive(row[indexes.concepto])).trim(),
      descripcion: String(cellToPrimitive(row[indexes.descripcion])).trim(),
      neto,
      sociedad: sociedad.trim() || 'CELSIA S.A.',
      ...(ceco ? { ceco } : {}),
      ...(empleadoRaw !== '' && empleadoRaw != null ? { empleado: empleadoRaw } : {}),
    })
  }

  return {
    rows,
    detailRows: rows.length,
    skippedTotalRows,
  }
}

export function relacionRowToArray(row: RelacionRow): (string | number)[] {
  return [
    row.anio,
    row.mes,
    row.periodo,
    row.cedula,
    row.nombreCompleto,
    row.concepto,
    row.descripcion,
    row.neto,
    row.sociedad,
  ]
}
