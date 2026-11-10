import type { CatalogoRow } from '../data/catalogos/types'
import type { RelacionRow } from './transformNomina'

export const PAGO_HEADERS = [
  'EMPRESA',
  'ACREEDOR',
  'CONCEPTO',
  'DESCRIPCION',
  'VALOR EN NOMINA',
  'Nº DOCUMENTO',
  'CLASE DE DOCUMENTO',
  'FECHA DE DOCUMENTO',
  'METODO DE PAGO',
  'VALOR CXP',
  'DIFERENCIA',
  'FECHA EN LA QUE SE RECIBE LA CXP',
  'FECHA DE ENVIO PARA PAGO',
  'SE ENVIA RELACION',
  'PERIODICIDAD',
] as const

/** Headers PAGO con CECO (Fundación Grupo Argos), tras Acreedor. */
export const PAGO_HEADERS_CON_CECO = [
  'EMPRESA',
  'ACREEDOR',
  'CECO',
  'CONCEPTO',
  'DESCRIPCION',
  'VALOR EN NOMINA',
  'Nº DOCUMENTO',
  'CLASE DE DOCUMENTO',
  'FECHA DE DOCUMENTO',
  'METODO DE PAGO',
  'VALOR CXP',
  'DIFERENCIA',
  'FECHA EN LA QUE SE RECIBE LA CXP',
  'FECHA DE ENVIO PARA PAGO',
  'SE ENVIA RELACION',
  'PERIODICIDAD',
] as const

export type PagoTercerosRow = {
  empresa: string
  acreedor: string
  ceco?: string
  concepto: string
  descripcion: string
  /** Número o '' (editable en blanco, p. ej. BRECHA Odinsa). */
  valorEnNomina: number | ''
  numeroDocumento: string
  claseDocumento: string
  fechaDocumento: string
  metodoPago: string
  valorCxp: string
  diferencia: string
  fechaRecibeCxp: string
  fechaEnvioPago: string
  seEnviaRelacion: string
  periodicidad: string
}

export type PagoTransformResult = {
  rows: PagoTercerosRow[]
  matchedConcepts: number
  catalogSize: number
}

function toNumber(value: number | string): number {
  if (typeof value === 'number') return value
  const n = Number(value)
  return Number.isFinite(n) ? n : 0
}

function isAcreedorNa(acreedor: string): boolean {
  const key = acreedor.trim().toUpperCase()
  return !key || key === 'N/A'
}

/** Acreedor A–Z (N/A al final); empate → concepto A–Z; empate → CECO A–Z. */
function comparePagoRows(a: PagoTercerosRow, b: PagoTercerosRow): number {
  const aNa = isAcreedorNa(a.acreedor)
  const bNa = isAcreedorNa(b.acreedor)
  if (aNa !== bNa) return aNa ? 1 : -1
  const byAcreedor = a.acreedor.trim().localeCompare(b.acreedor.trim(), 'es', {
    sensitivity: 'base',
  })
  if (byAcreedor !== 0) return byAcreedor
  const byConcepto = a.concepto.trim().localeCompare(b.concepto.trim(), 'es', {
    sensitivity: 'base',
  })
  if (byConcepto !== 0) return byConcepto
  return (a.ceco ?? '').trim().localeCompare((b.ceco ?? '').trim(), 'es', {
    sensitivity: 'base',
  })
}

function normalizeMetodoPago(metodo: string): string {
  return metodo
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toUpperCase()
}

export function isMetodoSoloSeEnviaRelacion(metodo: string): boolean {
  return normalizeMetodoPago(metodo) === 'SOLO SE ENVIA RELACION'
}

/** TRANSFERENCIA / typo TRANSFERANCIA / SOLO SE ENVÍA RELACIÓN (Celsia). */
export function isMetodoParaRelacionPorAcreedor(metodo: string): boolean {
  const n = normalizeMetodoPago(metodo)
  return n === 'TRANSFERENCIA' || n === 'TRANSFERANCIA' || n === 'SOLO SE ENVIA RELACION'
}

/** SUMMA: acreedor 30 (préstamos Grupo Argos / ORDEN DE PAGO) también va al ZIP. */
export function isSummaAcreedor30(row: PagoTercerosRow, grupoCatalogo: string): boolean {
  return (
    grupoCatalogo.trim().toUpperCase() === 'SUMMA' && row.acreedor.trim() === '30'
  )
}

/** Filas PAGO que generan Excel(s) en el ZIP de relaciones por acreedor. */
export function includeInRelacionPorAcreedor(
  row: PagoTercerosRow,
  grupoCatalogo: string,
): boolean {
  return (
    isMetodoParaRelacionPorAcreedor(row.metodoPago) ||
    isSummaAcreedor30(row, grupoCatalogo)
  )
}

/** @deprecated alias — usar isMetodoParaRelacionPorAcreedor */
export function isMetodoTransferencia(metodo: string): boolean {
  return isMetodoParaRelacionPorAcreedor(metodo)
}

function metodoPagoConsolidado(metodo: string): string {
  const n = normalizeMetodoPago(metodo)
  if (n === 'TRANSFERANCIA' || n === 'TRANSFERENCIA') return 'TRANSFERENCIA'
  if (n === 'SOLO SE ENVIA RELACION') return 'SOLO SE ENVÍA RELACIÓN'
  return metodo.trim()
}

/** Conceptos de PAGO incluidos en ZIP relaciones → filas RELACION que los usan. */
export function filterRelacionByTransferenciaConceptos(
  relacionRows: RelacionRow[],
  pagoRows: PagoTercerosRow[],
  grupoCatalogo = '',
): RelacionRow[] {
  const conceptos = new Set<string>()
  for (const row of pagoRows) {
    if (!includeInRelacionPorAcreedor(row, grupoCatalogo)) continue
    const concepto = row.concepto.trim().toUpperCase()
    if (concepto) conceptos.add(concepto)
  }
  if (conceptos.size === 0) return []
  return relacionRows.filter((row) => conceptos.has(row.concepto.trim().toUpperCase()))
}

const CXP_STRING_FIELDS = [
  'numeroDocumento',
  'claseDocumento',
  'fechaDocumento',
  'valorCxp',
  'diferencia',
  'fechaRecibeCxp',
  'fechaEnvioPago',
  'seEnviaRelacion',
  'periodicidad',
] as const

function mergeCxpField(values: string[]): string {
  const nonEmpty = values.map((v) => v.trim()).filter(Boolean)
  if (nonEmpty.length === 0) return ''
  const first = nonEmpty[0]
  return nonEmpty.every((v) => v === first) ? first : ''
}

/**
 * Suma duplicados exactos acreedor+concepto en transferencia y
 * SOLO SE ENVÍA RELACIÓN; no une conceptos distintos del mismo acreedor.
 * Otros métodos de pago se dejan sin consolidar.
 */
export function consolidateTransferenciaByAcreedorConcepto(
  rows: PagoTercerosRow[],
): PagoTercerosRow[] {
  const transfer: PagoTercerosRow[] = []
  const otros: PagoTercerosRow[] = []

  for (const row of rows) {
    if (isMetodoParaRelacionPorAcreedor(row.metodoPago)) transfer.push(row)
    else otros.push(row)
  }

  const groups = new Map<string, PagoTercerosRow[]>()
  for (const row of transfer) {
    const ceco = (row.ceco ?? '').trim()
    const key = `${row.acreedor.trim().toUpperCase()}|${row.concepto.trim().toUpperCase()}|${ceco}`
    const list = groups.get(key)
    if (list) list.push(row)
    else groups.set(key, [row])
  }

  const consolidated: PagoTercerosRow[] = []
  for (const group of groups.values()) {
    const first = group[0]
    const merged: PagoTercerosRow = {
      empresa: first.empresa,
      acreedor: first.acreedor,
      ...(first.ceco ? { ceco: first.ceco } : {}),
      concepto: first.concepto,
      descripcion: first.descripcion,
      valorEnNomina: group.reduce((sum, r) => sum + toNumber(r.valorEnNomina), 0),
      metodoPago: metodoPagoConsolidado(first.metodoPago),
      numeroDocumento: '',
      claseDocumento: '',
      fechaDocumento: '',
      valorCxp: '',
      diferencia: '',
      fechaRecibeCxp: '',
      fechaEnvioPago: '',
      seEnviaRelacion: '',
      periodicidad: '',
    }

    for (const field of CXP_STRING_FIELDS) {
      merged[field] = mergeCxpField(group.map((r) => r[field]))
    }

    consolidated.push(merged)
  }

  const result = [...otros, ...consolidated]
  result.sort(comparePagoRows)
  return result
}

const BRECHA_SOURCE = '4CBV'
const BRECHA_TARGET_SUMMA_CEMENTOS = 'BRECHA P'
const BRECHA_TARGET_GRUPO_ARGOS = 'BRECHA'
const BRECHA_TARGET_ODINSA = 'BRECHA'
/** OA01 / OG01 / OD01: BRECHA siempre, valor en blanco (no OV01). */
const ODINSA_BRECHA_EMPRESAS = new Set(['OA01', 'OG01', 'OD01'])
const SUMMA_IVA_SOURCE = '6CX1'
const SUMMA_IVA_TARGET = 'IVA'
const SUMMA_IVA_COMISION_POR_FILA = 12019

/** Cementos y filiales (no pensionados): conceptos con periodicidad MENSUAL. */
const PERIODICIDAD_MENSUAL_POR_EMPRESA = new Set([
  'co02',
  'co04',
  'logitrans',
  'casaparami',
  'solucionesmodulares',
])
const PERIODICIDAD_MENSUAL_CONCEPTOS = new Set(['4C43', '4C42', '6MB2'])

/**
 * Overrides de periodicidad por empresa interna + concepto.
 * Usa codigo de empresa (CO02, Logitrans…), no sociedad SAP.
 */
export function applyPeriodicidadOverridesByEmpresa(
  rows: PagoTercerosRow[],
  empresaCodigoInterno: string,
): PagoTercerosRow[] {
  const empresaKey = empresaCodigoInterno.trim().toLowerCase()
  if (!PERIODICIDAD_MENSUAL_POR_EMPRESA.has(empresaKey)) return rows

  return rows.map((row) => {
    const concepto = row.concepto.trim().toUpperCase()
    if (!PERIODICIDAD_MENSUAL_CONCEPTOS.has(concepto)) return row
    return { ...row, periodicidad: 'MENSUAL' }
  })
}

function formatFechaHoyLocal(): string {
  const now = new Date()
  const dd = String(now.getDate()).padStart(2, '0')
  const mm = String(now.getMonth() + 1).padStart(2, '0')
  const yyyy = now.getFullYear()
  return `${dd}/${mm}/${yyyy}`
}

/** Celsia Internet: clase AB siempre; si no hay fecha documento, usar fecha del día. */
export function applyCelsiaInternetPagoDefaults(
  rows: PagoTercerosRow[],
  empresaCodigoInterno: string,
): PagoTercerosRow[] {
  if (empresaCodigoInterno.trim().toLowerCase() !== 'celsia int') return rows
  const hoy = formatFechaHoyLocal()
  return rows.map((row) => {
    const fechaDocumento = row.fechaDocumento.trim() || hoy
    return {
      ...row,
      claseDocumento: 'AB',
      fechaDocumento,
      fechaRecibeCxp: fechaDocumento,
    }
  })
}

/** Cementos y filiales: clase de documento KO siempre. */
export function applyCementosPagoDefaults(
  rows: PagoTercerosRow[],
  grupoCatalogo: string,
): PagoTercerosRow[] {
  if (grupoCatalogo.trim().toUpperCase() !== 'CEMENTOS Y FILIALES') return rows
  return rows.map((row) => ({ ...row, claseDocumento: 'KO' }))
}

type ConceptoAggregate = {
  concepto: string
  descripcion: string
  total: number
  count: number
  ceco?: string
}

function buildPagoRow(
  empresaCodigo: string,
  cat: CatalogoRow,
  concepto: string,
  descripcion: string,
  valorEnNomina: number | '',
  periodicidadDefault: string,
  forcePeriodicidad: boolean,
  ceco?: string,
): PagoTercerosRow {
  return {
    empresa: empresaCodigo,
    acreedor: cat.acreedor || 'N/A',
    ...(ceco ? { ceco } : {}),
    concepto,
    descripcion: descripcion || cat.descripcion,
    valorEnNomina,
    numeroDocumento: '',
    claseDocumento: '',
    fechaDocumento: '',
    metodoPago: cat.metodoPago || 'N/A',
    valorCxp: '',
    diferencia: '',
    fechaRecibeCxp: '',
    fechaEnvioPago: '',
    seEnviaRelacion: '',
    periodicidad: forcePeriodicidad
      ? periodicidadDefault
      : cat.periodicidad?.trim() || periodicidadDefault,
  }
}

/**
 * Si viene 4CBV en nómina y no viene el concepto target (BRECHA P / BRECHA),
 * duplicar en PAGO con el valor de 4CBV y datos de catálogo del target.
 * Con split CECO: una fila target por cada CECO de 4CBV.
 */
function appendBrechaPensionalFrom4Cbv(
  rows: PagoTercerosRow[],
  aggregates: Map<string, ConceptoAggregate>,
  byConcepto: Map<string, CatalogoRow>,
  empresaCodigo: string,
  periodicidadDefault: string,
  forcePeriodicidad: boolean,
  targetConcepto: string,
): void {
  const targetKey = targetConcepto.trim().toUpperCase()
  const cat = byConcepto.get(targetKey)
  if (!cat) return

  const sources = [...aggregates.values()].filter(
    (agg) => agg.concepto.trim().toUpperCase() === BRECHA_SOURCE,
  )
  if (sources.length === 0) return

  for (const source of sources) {
    const ceco = source.ceco?.trim() || ''
    const alreadyInAgg = [...aggregates.values()].some(
      (agg) =>
        agg.concepto.trim().toUpperCase() === targetKey &&
        (agg.ceco?.trim() || '') === ceco,
    )
    if (alreadyInAgg) continue
    const alreadyInRows = rows.some(
      (row) =>
        row.concepto.trim().toUpperCase() === targetKey &&
        (row.ceco?.trim() || '') === ceco,
    )
    if (alreadyInRows) continue

    rows.push(
      buildPagoRow(
        empresaCodigo,
        cat,
        cat.concepto || targetConcepto,
        cat.descripcion || targetConcepto,
        source.total,
        periodicidadDefault,
        forcePeriodicidad,
        ceco || undefined,
      ),
    )
  }
}

/**
 * Odinsa OA01/OG01/OD01: siempre agregar BRECHA / CBP PLAN BRECHA
 * con valor en blanco (editable). No depende de 4CBV. No OV01.
 */
function appendOdinsaBrechaBlank(
  rows: PagoTercerosRow[],
  byConcepto: Map<string, CatalogoRow>,
  empresaCodigo: string,
  periodicidadDefault: string,
  forcePeriodicidad: boolean,
): void {
  const cat = byConcepto.get(BRECHA_TARGET_ODINSA)
  if (!cat) return
  const already = rows.some(
    (row) => row.concepto.trim().toUpperCase() === BRECHA_TARGET_ODINSA,
  )
  if (already) return

  rows.push(
    buildPagoRow(
      empresaCodigo,
      cat,
      cat.concepto || BRECHA_TARGET_ODINSA,
      cat.descripcion || 'CBP PLAN BRECHA',
      '',
      periodicidadDefault,
      forcePeriodicidad,
    ),
  )
}

/**
 * SUMMA / SATOR (1500): si hay filas 6CX1 en nómina y no viene IVA,
 * agregar en PAGO IVA = (# filas 6CX1) × 12019 con datos de catálogo.
 */
function appendIvaComisionEmbargos(
  rows: PagoTercerosRow[],
  aggregates: Map<string, ConceptoAggregate>,
  byConcepto: Map<string, CatalogoRow>,
  empresaCodigo: string,
  periodicidadDefault: string,
  forcePeriodicidad: boolean,
): void {
  const source = aggregates.get(SUMMA_IVA_SOURCE)
  if (!source || source.count <= 0) return
  if (aggregates.has(SUMMA_IVA_TARGET)) return
  const cat = byConcepto.get(SUMMA_IVA_TARGET)
  if (!cat) return

  rows.push(
    buildPagoRow(
      empresaCodigo,
      cat,
      cat.concepto || SUMMA_IVA_TARGET,
      cat.descripcion || 'EMBARGO CIVIL COMISION IVA',
      source.count * SUMMA_IVA_COMISION_POR_FILA,
      periodicidadDefault,
      forcePeriodicidad,
    ),
  )
}

export function transformPagoTerceros(
  relacionRows: RelacionRow[],
  empresaCodigo: string,
  catalogo: CatalogoRow[],
  periodicidadDefault = '',
  forcePeriodicidad = false,
  grupoCatalogo = '',
  /** Fundación 1600: agregar por concepto + CECO */
  splitByCeco = false,
): PagoTransformResult {
  const byConcepto = new Map<string, CatalogoRow>()
  for (const row of catalogo) {
    const key = row.concepto.trim().toUpperCase()
    if (!byConcepto.has(key)) byConcepto.set(key, row)
  }

  const aggregates = new Map<string, ConceptoAggregate>()

  for (const row of relacionRows) {
    const concepto = row.concepto.trim()
    if (!concepto) continue
    const ceco = splitByCeco ? (row.ceco?.trim() || '') : ''
    const key = splitByCeco
      ? `${concepto.toUpperCase()}|${ceco}`
      : concepto.toUpperCase()
    const current = aggregates.get(key)
    if (current) {
      current.total += toNumber(row.neto)
      current.count += 1
    } else {
      aggregates.set(key, {
        concepto,
        descripcion: row.descripcion.trim(),
        total: toNumber(row.neto),
        count: 1,
        ...(ceco ? { ceco } : {}),
      })
    }
  }

  const rows: PagoTercerosRow[] = []

  for (const [, agg] of aggregates) {
    const cat = byConcepto.get(agg.concepto.trim().toUpperCase())
    if (!cat) continue

    rows.push(
      buildPagoRow(
        empresaCodigo,
        cat,
        agg.concepto,
        agg.descripcion || cat.descripcion,
        agg.total,
        periodicidadDefault,
        forcePeriodicidad,
        splitByCeco ? agg.ceco : undefined,
      ),
    )
  }

  const grupo = grupoCatalogo.trim().toUpperCase()
  const brechaTarget =
    grupo === 'SUMMA' || grupo === 'CEMENTOS Y FILIALES'
      ? BRECHA_TARGET_SUMMA_CEMENTOS
      : grupo === 'GRUPO ARGOS Y FILIALES'
        ? BRECHA_TARGET_GRUPO_ARGOS
        : ''
  if (brechaTarget) {
    appendBrechaPensionalFrom4Cbv(
      rows,
      aggregates,
      byConcepto,
      empresaCodigo,
      periodicidadDefault,
      forcePeriodicidad,
      brechaTarget,
    )
  }
  if (ODINSA_BRECHA_EMPRESAS.has(empresaCodigo.trim().toUpperCase())) {
    appendOdinsaBrechaBlank(
      rows,
      byConcepto,
      empresaCodigo,
      periodicidadDefault,
      forcePeriodicidad,
    )
  }
  if (grupo === 'SUMMA' || empresaCodigo.trim() === '1500') {
    appendIvaComisionEmbargos(
      rows,
      aggregates,
      byConcepto,
      empresaCodigo,
      periodicidadDefault,
      forcePeriodicidad,
    )
  }

  rows.sort(comparePagoRows)

  return {
    rows,
    matchedConcepts: rows.length,
    catalogSize: byConcepto.size,
  }
}

export function pagoRowToArray(
  row: PagoTercerosRow,
  includeCeco = false,
): (string | number)[] {
  if (includeCeco) {
    return [
      row.empresa,
      row.acreedor,
      row.ceco ?? '',
      row.concepto,
      row.descripcion,
      row.valorEnNomina,
      row.numeroDocumento,
      row.claseDocumento,
      row.fechaDocumento,
      row.metodoPago,
      row.valorCxp,
      row.diferencia,
      row.fechaRecibeCxp,
      row.fechaEnvioPago,
      row.seEnviaRelacion,
      row.periodicidad,
    ]
  }
  return [
    row.empresa,
    row.acreedor,
    row.concepto,
    row.descripcion,
    row.valorEnNomina,
    row.numeroDocumento,
    row.claseDocumento,
    row.fechaDocumento,
    row.metodoPago,
    row.valorCxp,
    row.diferencia,
    row.fechaRecibeCxp,
    row.fechaEnvioPago,
    row.seEnviaRelacion,
    row.periodicidad,
  ]
}
