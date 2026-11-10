import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { getCatalogoByGrupo } from '../data/catalogos'
import { getPaletteForEmpresa, type BrandPalette } from '../data/brandPalettes'
import { codigoEnInforme, getEmpresaByCodigo, type EmpresaSap } from '../data/empresas'
import {
  RELACION_HEADERS,
  detectPeriodoFromNominaMatrix,
  relacionRowToArray,
  transformNominaRows,
  type RelacionRow,
} from './transformNomina'
import {
  PAGO_HEADERS,
  PAGO_HEADERS_CON_CECO,
  applyCementosPagoDefaults,
  applyCelsiaInternetPagoDefaults,
  applyPeriodicidadOverridesByEmpresa,
  consolidateTransferenciaByAcreedorConcepto,
  includeInRelacionPorAcreedor,
  isMetodoSoloSeEnviaRelacion,
  pagoRowToArray,
  transformPagoTerceros,
  type PagoTercerosRow,
} from './transformPagoTerceros'
import { CXP_SHEET_NAME, enrichPagoFromCxp, type EnrichCxpResult } from './transformCxp'
import {
  contingenteVariantFromGrupo,
  empresaUsaContingenteOrganizado,
  writeOrganizedContingenteSheet,
  type ContingenteVariant,
} from './transformContingente'
import {
  buildAssignmentCostCenterMap,
  buildCecoLabelByPersonal,
  enrichNominaWithCentroCostos,
  empresaUsaCentroCostosMaster,
} from './enrichNominaCentroCostos'
import {
  FIA_CONCEPTOS,
  FIA_INF_SHEET_NAME,
  buildFiaInfConceptoFromRelacionRows,
  empresaUsaFiaInfConcepto,
  isFiaAcreedor,
  type FiaInfResult,
} from './transformFiaInfConcepto'
import {
  GRANCOOP_TD_CONCEPTOS,
  buildGrancoopBundle,
  empresaUsaGrancoopCentralizador,
  isGrancoopAcreedor,
  parseCentralizadorNovedades,
  type CentralizadorNovedad,
  type GrancoopBundle,
} from './transformGrancoop'

export const NOMINA_SHEET_NAME = 'NOMINA'
export const CONTINGENTE_SHEET_NAME = 'CONTINGENTE'
export const RELACION_SHEET_NAME = 'RELACION TERCEROS'
export const PAGO_SHEET_NAME = 'PAGO TERCEROS'
export { CXP_SHEET_NAME }

const RELACION_LABELS: Record<(typeof RELACION_HEADERS)[number], string> = {
  AÑO: 'Año',
  MES: 'Mes',
  PERIODO: 'Periodo',
  CEDULA: 'Cédula',
  NOMBRE_COMPLETO: 'Nombre completo',
  CONCEPTO: 'Concepto',
  DESCRIPCION: 'Descripción',
  NETO: 'Neto',
  SOCIEDAD: 'Sociedad',
}

const PAGO_LABELS: Record<string, string> = {
  EMPRESA: 'Empresa',
  ACREEDOR: 'Acreedor',
  CECO: 'CECO',
  CONCEPTO: 'Concepto',
  DESCRIPCION: 'Descripción',
  'VALOR EN NOMINA': 'Valor en nómina',
  'Nº DOCUMENTO': 'Nº documento',
  'CLASE DE DOCUMENTO': 'Clase de documento',
  'FECHA DE DOCUMENTO': 'Fecha de documento',
  'METODO DE PAGO': 'Método de pago',
  'VALOR CXP': 'Valor CXP',
  DIFERENCIA: 'Diferencia',
  'FECHA EN LA QUE SE RECIBE LA CXP': 'Fecha recibe CXP',
  'FECHA DE ENVIO PARA PAGO': 'Fecha envío pago',
  'SE ENVIA RELACION': 'Se envía relación',
  PERIODICIDAD: 'Periodicidad',
}

const RELACION_WIDTHS = [8, 8, 10, 14, 38, 12, 34, 14, 18]
const PAGO_WIDTHS = [10, 12, 12, 32, 16, 14, 14, 14, 18, 12, 12, 16, 16, 14, 12]
const PAGO_WIDTHS_CON_CECO = [10, 12, 12, 12, 32, 16, 14, 14, 14, 18, 12, 12, 16, 16, 14, 12]

export type CxpEnrichSummary = {
  fileName: string
  updatedRows: number
  rowsWithDocumento: number
  cxpRowsUsed: number
  matchedByAcreedor: number
  filledWithCompanyDefault: number
  usedSociedadFallback?: boolean
  missingHeaders?: string[]
  error?: string
}

export type ContingenteSummary = {
  fileName: string
  dataRows: number
  error?: string
}

export type CentroCostosSummary = {
  fileName: string
  matched: number
  totalDataRows: number
  error?: string
}

export type ProcessSummary = {
  fileName: string
  empresaCodigo: string
  empresaNombre: string
  grupoCatalogo: string
  periodoResumen: string
  relacionRows: number
  pagoRows: number
  skippedTotalRows: number
  catalogSize: number
  outputSheets: string[]
  cxpEnrichment?: CxpEnrichSummary | null
  contingente?: ContingenteSummary | null
  centroCostos?: CentroCostosSummary | null
}

export type ProcessOutput = {
  workbook: ExcelJS.Workbook
  relacionRows: RelacionRow[]
  pagoRows: PagoTercerosRow[]
  summary: ProcessSummary
}

function findSheet(workbook: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
  const target = name.trim().toLowerCase()
  return workbook.worksheets.find((ws) => ws.name.trim().toLowerCase() === target)
}

function normalizeHeaderCell(value: unknown): string {
  if (value == null) return ''
  let text = ''
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    text = String(value)
  } else if (value instanceof Date) {
    text = value.toISOString()
  } else if (typeof value === 'object') {
    const obj = value as { text?: unknown; result?: unknown; richText?: Array<{ text?: string }> }
    if (typeof obj.text === 'string') text = obj.text
    else if (obj.result != null) return normalizeHeaderCell(obj.result)
    else if (Array.isArray(obj.richText)) text = obj.richText.map((p) => p.text ?? '').join('')
    else text = String(value)
  } else {
    text = String(value)
  }
  return text
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[º°]/g, '')
    .replace(/\./g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

/** CXP SAP, Internet (Doc/Tercero) o contable (NIT + Consecutivo/Tipo Doc/Fecha). */
function sheetLooksLikeCxp(worksheet: ExcelJS.Worksheet): boolean {
  const headerRow = worksheet.getRow(1)
  const headers: string[] = []
  headerRow.eachCell({ includeEmpty: true }, (cell, col) => {
    headers[col - 1] = normalizeHeaderCell(cell.value)
  })
  const set = new Set(headers.filter(Boolean))
  const hasSap = set.has('sociedad') || set.has('proveedor')
  const hasInternet = set.has('doc') && set.has('tercero')
  const hasContableId = set.has('nit') || set.has('proveedor')
  const hasContableDoc =
    set.has('consecutivo') || set.has('tipo doc') || set.has('fecha')
  const hasContable = hasContableId && hasContableDoc
  return hasSap || hasInternet || hasContable
}

/**
 * Prefiere pestaña "CXP". Si no existe, primera hoja con layout CXP
 * (SAP o Doc/Tercero de Celsia Internet). En archivo CXP dedicado, ultima opción: primera hoja.
 */
function findCxpSheet(
  workbook: ExcelJS.Workbook,
  options?: { allowFirstSheetFallback?: boolean },
): ExcelJS.Worksheet | undefined {
  const named = findSheet(workbook, CXP_SHEET_NAME)
  if (named) return named

  for (const ws of workbook.worksheets) {
    if (ws.name.trim().toLowerCase() === NOMINA_SHEET_NAME.toLowerCase()) continue
    if (sheetLooksLikeCxp(ws)) return ws
  }

  if (options?.allowFirstSheetFallback && workbook.worksheets.length > 0) {
    return workbook.worksheets[0]
  }
  return undefined
}

function worksheetToMatrix(worksheet: ExcelJS.Worksheet): unknown[][] {
  const matrix: unknown[][] = []
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const values = row.values as unknown[]
    matrix.push(values.slice(1))
  })
  return matrix
}

function thinBorder(palette: BrandPalette): Partial<ExcelJS.Borders> {
  const edge: Partial<ExcelJS.Border> = {
    style: 'thin',
    color: { argb: `FF${palette.border}` },
  }
  return { top: edge, left: edge, bottom: edge, right: edge }
}

function accentBorder(palette: BrandPalette): Partial<ExcelJS.Borders> {
  const edge: Partial<ExcelJS.Border> = {
    style: 'thin',
    color: { argb: `FF${palette.metaAccent}` },
  }
  return { top: edge, left: edge, bottom: edge, right: edge }
}

function colLetter(colCount: number): string {
  let n = colCount
  let s = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    s = String.fromCharCode(65 + rem) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

function labelPeriodoResumen(periodo: string | number): string {
  const key = String(periodo).trim()
  if (key === '10') return '1ª quincena'
  if (key === '20') return '2ª quincena'
  return `Periodo ${periodo}`
}

function buildPeriodoResumen(rows: RelacionRow[]): string {
  if (rows.length === 0) return 'Sin datos'
  const first = rows[0]
  return `${String(first.mes).padStart(2, '0')}/${first.anio} · ${labelPeriodoResumen(first.periodo)} (${first.periodo})`
}

function styleHeaderRow(
  sheet: ExcelJS.Worksheet,
  rowNumber: number,
  labels: string[],
  palette: BrandPalette,
) {
  const headerRow = sheet.getRow(rowNumber)
  headerRow.height = 24
  labels.forEach((label, index) => {
    const cell = headerRow.getCell(index + 1)
    cell.value = label
    cell.font = {
      name: 'Calibri',
      size: 11,
      bold: true,
      color: { argb: `FF${palette.headerFg}` },
    }
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${palette.headerBg}` },
    }
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
    cell.border = thinBorder(palette)
  })
}

function styleTitleAndMeta(
  sheet: ExcelJS.Worksheet,
  colCount: number,
  titleText: string,
  metaText: string | null,
  palette: BrandPalette,
) {
  const lastCol = colLetter(colCount)
  sheet.mergeCells(`A1:${lastCol}1`)
  const title = sheet.getCell('A1')
  title.value = titleText
  title.font = { name: 'Calibri', size: 16, bold: true, color: { argb: `FF${palette.titleFg}` } }
  title.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: `FF${palette.titleBg}` },
  }
  title.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  sheet.getRow(1).height = 32

  if (metaText == null) return

  sheet.mergeCells(`A2:${lastCol}2`)
  const meta = sheet.getCell('A2')
  meta.value = metaText
  meta.font = { name: 'Calibri', size: 11, color: { argb: `FF${palette.metaFg}` } }
  meta.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: `FF${palette.metaBg}` },
  }
  meta.border = accentBorder(palette)
  meta.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  sheet.getRow(2).height = 22
}

function styleDataRow(
  excelRow: ExcelJS.Row,
  zebra: boolean,
  palette: BrandPalette,
  options?: { moneyCols?: number[]; leftCols?: number[] },
) {
  const moneyCols = new Set(options?.moneyCols ?? [])
  const leftCols = new Set(options?.leftCols ?? [])
  excelRow.height = 18
  excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.font = { name: 'Calibri', size: 10, color: { argb: `FF${palette.ink}` } }
    cell.border = thinBorder(palette)
    cell.alignment = {
      vertical: 'middle',
      horizontal: leftCols.has(colNumber) ? 'left' : 'center',
    }
    if (zebra) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${palette.zebra}` },
      }
    }
    if (moneyCols.has(colNumber)) {
      cell.numFmt = '#,##0'
      cell.alignment = { vertical: 'middle', horizontal: 'right' }
    }
  })
}

function toNetoNumber(value: number | string): number {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const n = Number(String(value).replace(/,/g, '').trim())
  return Number.isFinite(n) ? n : 0
}

function formatTotalMeta(total: number): string {
  return total.toLocaleString('es-CO', { maximumFractionDigits: 0 })
}

function addRelacionSheet(
  workbook: ExcelJS.Workbook,
  rows: RelacionRow[],
  empresa: EmpresaSap,
  periodoTexto: string,
  palette: BrandPalette,
  options?: {
    sheetName?: string
    titleText?: string
    metaExtra?: string
    omitMeta?: boolean
  },
) {
  const sheetName = options?.sheetName ?? RELACION_SHEET_NAME
  const titleText = options?.titleText ?? 'RELACIÓN DE TERCEROS'
  const omitMeta = Boolean(options?.omitMeta)
  const metaExtra = options?.metaExtra ? `    |    ${options.metaExtra}` : ''
  const totalNeto = rows.reduce((sum, row) => sum + toNetoNumber(row.neto), 0)
  const headerRow = omitMeta ? 2 : 3

  const sheet = workbook.addWorksheet(sheetName, {
    views: [{ state: 'frozen', ySplit: headerRow, showGridLines: false }],
    properties: { defaultRowHeight: 18 },
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  })

  const colCount = RELACION_HEADERS.length
  styleTitleAndMeta(
    sheet,
    colCount,
    titleText,
    omitMeta
      ? null
      : `Empresa: ${codigoEnInforme(empresa)} — ${empresa.nombre}    |    Periodo: ${periodoTexto}    |    Registros: ${rows.length}    |    Total: $${formatTotalMeta(totalNeto)}${metaExtra}`,
    palette,
  )
  styleHeaderRow(
    sheet,
    headerRow,
    RELACION_HEADERS.map((key) => RELACION_LABELS[key]),
    palette,
  )

  rows.forEach((row, index) => {
    const excelRow = sheet.addRow(relacionRowToArray(row))
    styleDataRow(excelRow, index % 2 === 1, palette, { moneyCols: [8], leftCols: [5, 7, 9] })
  })

  const totalRow = sheet.addRow(['', '', '', '', '', '', 'TOTAL', totalNeto, ''])
  totalRow.height = 20
  totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: `FF${palette.ink}` },
    }
    cell.border = thinBorder(palette)
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${palette.metaBg}` },
    }
    cell.alignment = {
      vertical: 'middle',
      horizontal: colNumber === 7 ? 'left' : colNumber === 8 ? 'right' : 'center',
    }
    if (colNumber === 8) {
      cell.numFmt = '#,##0'
    }
  })

  RELACION_WIDTHS.forEach((width, idx) => {
    sheet.getColumn(idx + 1).width = width
  })

  sheet.autoFilter = {
    from: { row: headerRow, column: 1 },
    to: { row: Math.max(headerRow, headerRow + rows.length), column: colCount },
  }
}

function addPagoSheet(
  workbook: ExcelJS.Workbook,
  rows: PagoTercerosRow[],
  empresa: EmpresaSap,
  palette: BrandPalette,
) {
  const totalValorNomina = rows.reduce((sum, row) => sum + toNetoNumber(row.valorEnNomina), 0)
  const includeCeco = empresaUsaCentroCostosMaster(empresa.codigo)
  const headers = includeCeco ? [...PAGO_HEADERS_CON_CECO] : [...PAGO_HEADERS]
  const widths = includeCeco ? PAGO_WIDTHS_CON_CECO : PAGO_WIDTHS
  const moneyCol = includeCeco ? 6 : 5
  const descCol = includeCeco ? 5 : 4
  const metodoCol = includeCeco ? 10 : 9

  const sheet = workbook.addWorksheet(PAGO_SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 3, showGridLines: false }],
    properties: { defaultRowHeight: 18 },
    pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 },
  })

  const colCount = headers.length
  styleTitleAndMeta(
    sheet,
    colCount,
    'PAGO A TERCEROS',
    `Empresa: ${codigoEnInforme(empresa)} — ${empresa.nombre}    |    Catálogo: ${empresa.grupoCatalogo}    |    Registros: ${rows.length}    |    Total: $${formatTotalMeta(totalValorNomina)}`,
    palette,
  )
  styleHeaderRow(
    sheet,
    3,
    headers.map((key) => PAGO_LABELS[key] ?? key),
    palette,
  )

  rows.forEach((row, index) => {
    const excelRow = sheet.addRow(pagoRowToArray(row, includeCeco))
    styleDataRow(excelRow, index % 2 === 1, palette, {
      moneyCols: [moneyCol],
      leftCols: [descCol, metodoCol],
    })
  })

  const totalValues = Array.from({ length: colCount }, () => '' as string | number)
  totalValues[descCol - 1] = 'TOTAL'
  totalValues[moneyCol - 1] = totalValorNomina
  const totalRow = sheet.addRow(totalValues)
  totalRow.height = 20
  totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: `FF${palette.ink}` },
    }
    cell.border = thinBorder(palette)
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${palette.metaBg}` },
    }
    cell.alignment = {
      vertical: 'middle',
      horizontal:
        colNumber === descCol ? 'left' : colNumber === moneyCol ? 'right' : 'center',
    }
    if (colNumber === moneyCol) {
      cell.numFmt = '#,##0'
    }
  })

  widths.forEach((width, idx) => {
    sheet.getColumn(idx + 1).width = width
  })

  sheet.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: Math.max(3, 3 + rows.length), column: colCount },
  }

  if (includeCeco) {
    let sumFundacion = 0
    let sumGanaderia = 0
    for (const row of rows) {
      const ceco = (row.ceco ?? '').trim().toUpperCase()
      const valor = toNetoNumber(row.valorEnNomina)
      if (ceco === 'FUNDACION') sumFundacion += valor
      else if (ceco === 'GANADERIA') sumGanaderia += valor
    }
    const sumCecoTotal = sumFundacion + sumGanaderia
    const boxStartRow = totalRow.number + 2
    const moneyFmt = '"$"#,##0'

    const styleCecoLabel = (cell: ExcelJS.Cell) => {
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: `FF${palette.ink}` },
      }
      cell.border = thinBorder(palette)
      cell.alignment = { vertical: 'middle', horizontal: 'left' }
    }
    const styleCecoValue = (cell: ExcelJS.Cell, yellow = false) => {
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: yellow,
        color: { argb: `FF${palette.ink}` },
      }
      cell.border = thinBorder(palette)
      cell.alignment = { vertical: 'middle', horizontal: 'right' }
      cell.numFmt = moneyFmt
      if (yellow) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFFFFF00' },
        }
      }
    }

    const r1 = sheet.getRow(boxStartRow)
    r1.getCell(1).value = 'FUNDACION'
    r1.getCell(2).value = sumFundacion
    styleCecoLabel(r1.getCell(1))
    styleCecoValue(r1.getCell(2))

    const r2 = sheet.getRow(boxStartRow + 1)
    r2.getCell(1).value = 'GANADERIA'
    r2.getCell(2).value = sumGanaderia
    styleCecoLabel(r2.getCell(1))
    styleCecoValue(r2.getCell(2))

    const r3 = sheet.getRow(boxStartRow + 2)
    r3.getCell(1).value = ''
    r3.getCell(2).value = sumCecoTotal
    styleCecoLabel(r3.getCell(1))
    styleCecoValue(r3.getCell(2), true)
  }
}

/** Copia una hoja al workbook de salida (valores, anchos, merges). */
function copySheet(
  target: ExcelJS.Workbook,
  source: ExcelJS.Worksheet,
  sheetName: string,
): void {
  const dest = target.addWorksheet(sheetName, {
    views: source.views,
    properties: source.properties,
  })

  source.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const newRow = dest.getRow(rowNumber)
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber)
      newCell.value = cell.value
      if (cell.style) {
        newCell.style = { ...cell.style }
      }
      if (cell.numFmt) newCell.numFmt = cell.numFmt
    })
    if (row.height) newRow.height = row.height
  })

  source.columns.forEach((col, idx) => {
    if (col && col.width != null) {
      dest.getColumn(idx + 1).width = col.width
    }
  })

  const merges = (source as ExcelJS.Worksheet & { model?: { merges?: string[] } }).model?.merges
  if (Array.isArray(merges)) {
    for (const merge of merges) {
      try {
        dest.mergeCells(merge)
      } catch {
        // ignorar merges inválidos al copiar
      }
    }
  }
}

function buildOutputSheetNames(hasNomina: boolean, _hasCxp = true): string[] {
  const sheets: string[] = []
  if (hasNomina) sheets.push(NOMINA_SHEET_NAME)
  // CXP siempre en salida (copia o vacía)
  sheets.push(CXP_SHEET_NAME)
  sheets.push(CONTINGENTE_SHEET_NAME, RELACION_SHEET_NAME, PAGO_SHEET_NAME)
  return sheets
}

/** Pestaña CXP vacía (solo encabezados) para rellenar después. */
function addBlankCxpSheet(workbook: ExcelJS.Workbook): ExcelJS.Worksheet {
  const sheet = workbook.addWorksheet(CXP_SHEET_NAME)
  const headers = [
    'Sociedad',
    'Proveedor',
    'Nº documento',
    'Clase de documento',
    'Fe.contabilización',
  ]
  headers.forEach((label, index) => {
    sheet.getCell(1, index + 1).value = label
  })
  sheet.getRow(1).font = { bold: true }
  return sheet
}

function organizeContingenteSheet(
  source: ExcelJS.Worksheet,
  variant: ContingenteVariant,
): {
  sheet: ExcelJS.Worksheet
  dataRows: number
} {
  const tmp = new ExcelJS.Workbook()
  return writeOrganizedContingenteSheet(tmp, source, CONTINGENTE_SHEET_NAME, variant)
}

function findContingenteSheet(
  workbook: ExcelJS.Workbook,
  options?: { allowFirstSheetFallback?: boolean },
): ExcelJS.Worksheet | undefined {
  const named = findSheet(workbook, CONTINGENTE_SHEET_NAME)
  if (named) return named
  if (options?.allowFirstSheetFallback && workbook.worksheets.length > 0) {
    return workbook.worksheets[0]
  }
  return undefined
}

function addGrancoopSheets(
  workbook: ExcelJS.Workbook,
  bundle: GrancoopBundle,
  palette: BrandPalette,
) {
  // --- TD X CONCEPTO ---
  {
    const sheet = workbook.addWorksheet('TD X CONCEPTO', {
      views: [{ state: 'frozen', ySplit: 4, showGridLines: false }],
    })
    sheet.getRow(3).getCell(1).value = 'Suma de DESCUENTOS'
    sheet.getRow(3).getCell(3).value = 'CONCEPTO'
    sheet.getRow(3).font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: `FF${palette.ink}` },
    }
    const headers = [
      'CEDULA',
      'NOMBRE_COMPLETO',
      ...GRANCOOP_TD_CONCEPTOS,
      'Total general',
    ]
    styleHeaderRow(sheet, 4, headers, palette)
    const moneyCols = Array.from({ length: GRANCOOP_TD_CONCEPTOS.length + 1 }, (_, i) => i + 3)
    bundle.td.forEach((row, index) => {
      const excelRow = sheet.addRow([
        row.cedula,
        row.nombreCompleto,
        ...GRANCOOP_TD_CONCEPTOS.map((c) => row.valores[c] ?? ''),
        row.total,
      ])
      styleDataRow(excelRow, index % 2 === 1, palette, { moneyCols, leftCols: [2] })
    })
    const totalRow = sheet.addRow([
      'Total general',
      '',
      ...GRANCOOP_TD_CONCEPTOS.map((c) => bundle.tdColumnTotals[c]),
      bundle.tdGrandTotal,
    ])
    totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: `FF${palette.ink}` },
      }
      cell.border = thinBorder(palette)
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${palette.metaBg}` },
      }
      if (colNumber >= 3) {
        cell.numFmt = '#,##0'
        cell.alignment = { vertical: 'middle', horizontal: 'right' }
      }
    })
    ;[12, 36, ...GRANCOOP_TD_CONCEPTOS.map(() => 12), 14].forEach((w, i) => {
      sheet.getColumn(i + 1).width = w
    })
  }

  // --- NOMINA ---
  {
    const sheet = workbook.addWorksheet('NOMINA', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    })
    styleHeaderRow(
      sheet,
      1,
      [
        'EMPLEADO',
        'CEDULA',
        'NOMBRE_COMPLETO',
        'CONCEPTO',
        'DESCRIPCION',
        'DESCUENTOS',
        'SOCIEDAD',
      ],
      palette,
    )
    bundle.nomina.forEach((row, index) => {
      const excelRow = sheet.addRow([
        row.empleado,
        row.cedula,
        row.nombreCompleto,
        row.concepto,
        row.descripcion,
        row.descuentos,
        row.sociedad,
      ])
      styleDataRow(excelRow, index % 2 === 1, palette, {
        moneyCols: [6],
        leftCols: [3, 5, 7],
      })
    })
    ;[12, 14, 36, 10, 28, 14, 28].forEach((w, i) => {
      sheet.getColumn(i + 1).width = w
    })
  }

  // --- INFORME FINAL ---
  {
    const sheet = workbook.addWorksheet('INFORME FINAL', {
      views: [{ state: 'frozen', ySplit: 3, showGridLines: false }],
    })
    styleHeaderRow(
      sheet,
      3,
      [
        'Empleado',
        'Colaborador',
        'Concepto',
        'Suma de Valor',
        'Valor en nomina',
        'Diferencia',
        'Observaciones',
      ],
      palette,
    )
    bundle.informe.forEach((row, index) => {
      const excelRow = sheet.addRow([
        row.empleado,
        row.colaborador,
        row.concepto,
        row.sumaValor,
        row.valorEnNomina,
        row.diferencia,
        row.observaciones,
      ])
      styleDataRow(excelRow, index % 2 === 1, palette, {
        moneyCols: [4, 5, 6],
        leftCols: [2, 7],
      })
    })
    ;[12, 36, 10, 14, 14, 12, 32].forEach((w, i) => {
      sheet.getColumn(i + 1).width = w
    })
  }

  // --- CARTERA ---
  {
    const sheet = workbook.addWorksheet('CARTERA', {
      views: [{ state: 'frozen', ySplit: 1, showGridLines: false }],
    })
    styleHeaderRow(sheet, 1, ['CEDULA', 'NOMBRE', 'CONCEPTO', 'DESCUENTO'], palette)
    bundle.cartera.forEach((row, index) => {
      const excelRow = sheet.addRow([
        row.cedula,
        row.nombre,
        row.concepto,
        row.descuento,
      ])
      styleDataRow(excelRow, index % 2 === 1, palette, {
        moneyCols: [4],
        leftCols: [2],
      })
    })
    const totalDescuento = bundle.cartera.reduce((sum, row) => sum + row.descuento, 0)
    const totalRow = sheet.addRow(['', '', 'TOTAL', totalDescuento])
    totalRow.height = 20
    totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font = {
        name: 'Calibri',
        size: 10,
        bold: true,
        color: { argb: `FF${palette.ink}` },
      }
      cell.border = thinBorder(palette)
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: `FF${palette.metaBg}` },
      }
      cell.alignment = {
        vertical: 'middle',
        horizontal: colNumber === 4 ? 'right' : colNumber === 3 ? 'left' : 'center',
      }
      if (colNumber === 4) cell.numFmt = '#,##0'
    })
    ;[14, 36, 10, 14].forEach((w, i) => {
      sheet.getColumn(i + 1).width = w
    })
  }
}

function addFiaInfSheet(
  workbook: ExcelJS.Workbook,
  fia: FiaInfResult,
  palette: BrandPalette,
) {
  const colCount = 2 + FIA_CONCEPTOS.length + 1
  const sheet = workbook.addWorksheet(FIA_INF_SHEET_NAME, {
    views: [{ state: 'frozen', ySplit: 2, showGridLines: false }],
    properties: { defaultRowHeight: 18 },
  })

  // R1: título estilo pivot del ejemplo
  const titleRow = sheet.getRow(1)
  titleRow.getCell(1).value = 'Suma de TOTAL'
  titleRow.getCell(3).value = 'CONCEPTO'
  titleRow.font = { name: 'Calibri', size: 10, bold: true, color: { argb: `FF${palette.ink}` } }

  const headerLabels = [
    'NOMBRE EMPLEADO',
    'N° DOCUMENTO',
    ...FIA_CONCEPTOS,
    'Total general',
  ]
  styleHeaderRow(sheet, 2, headerLabels, palette)

  const moneyCols = Array.from({ length: FIA_CONCEPTOS.length + 1 }, (_, i) => i + 3)

  fia.rows.forEach((row, index) => {
    const values: (string | number)[] = [
      row.nombre,
      row.documento,
      ...FIA_CONCEPTOS.map((c) => row.valores[c] ?? ''),
      row.total,
    ]
    const excelRow = sheet.addRow(values)
    styleDataRow(excelRow, index % 2 === 1, palette, {
      moneyCols,
      leftCols: [1],
    })
  })

  const totalValues: (string | number)[] = [
    'Total general',
    '',
    ...FIA_CONCEPTOS.map((c) => fia.columnTotals[c]),
    fia.grandTotal,
  ]
  const totalRow = sheet.addRow(totalValues)
  totalRow.height = 20
  totalRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    cell.font = {
      name: 'Calibri',
      size: 10,
      bold: true,
      color: { argb: `FF${palette.ink}` },
    }
    cell.border = thinBorder(palette)
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: `FF${palette.metaBg}` },
    }
    cell.alignment = {
      vertical: 'middle',
      horizontal: colNumber === 1 ? 'left' : colNumber >= 3 ? 'right' : 'center',
    }
    if (colNumber >= 3) cell.numFmt = '#,##0'
  })

  const widths = [36, 14, ...FIA_CONCEPTOS.map(() => 12), 14]
  widths.forEach((width, idx) => {
    sheet.getColumn(idx + 1).width = width
  })

  sheet.autoFilter = {
    from: { row: 2, column: 1 },
    to: { row: Math.max(2, 2 + fia.rows.length), column: colCount },
  }
}

function createOutputWorkbook(
  relacionRows: RelacionRow[],
  pagoRows: PagoTercerosRow[],
  empresa: EmpresaSap,
  nominaSheet?: ExcelJS.Worksheet,
  cxpSheet?: ExcelJS.Worksheet,
  contingenteSheet?: ExcelJS.Worksheet,
): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'NOMINA → RELACION / PAGO TERCEROS'
  workbook.created = new Date()

  if (nominaSheet) {
    copySheet(workbook, nominaSheet, NOMINA_SHEET_NAME)
  }
  if (cxpSheet) {
    copySheet(workbook, cxpSheet, CXP_SHEET_NAME)
  } else {
    addBlankCxpSheet(workbook)
  }
  if (contingenteSheet) {
    copySheet(workbook, contingenteSheet, CONTINGENTE_SHEET_NAME)
  } else {
    workbook.addWorksheet(CONTINGENTE_SHEET_NAME)
  }

  const palette = getPaletteForEmpresa(empresa)
  const periodoTexto = buildPeriodoResumen(relacionRows)
  addRelacionSheet(workbook, relacionRows, empresa, periodoTexto, palette)
  addPagoSheet(workbook, pagoRows, empresa, palette)

  return workbook
}

const MESES_ES = [
  'enero',
  'febrero',
  'marzo',
  'abril',
  'mayo',
  'junio',
  'julio',
  'agosto',
  'septiembre',
  'octubre',
  'noviembre',
  'diciembre',
] as const

/** Prefijos de descripción de concepto → nombre legible del acreedor. */
const DESCRIPCION_PREFIXES = [
  /^APOR\.?\s*FDO\.?\s*/i,
  /^APOR\.?\s*FEDE\.?\s*/i,
  /^AHORRO\s+FDO\.?\s*/i,
  /^CUOTA\s+COOP\.?\s*/i,
  /^PTMO\.?\s*/i,
  /^APOR\.?\s*/i,
  /^LIBR\.?\s*/i,
  /^CUOTA\s+/i,
  /^DCTO\.?\s*/i,
  /^APV\.?\s*/i,
  /^VIVIENDA\.?\s*/i,
]

const EMPRESA_NOMBRE_ARCHIVO: Record<string, string> = {
  CO01: 'Celsia SA',
  CO05: 'Celsia Colombia',
  CETSA: 'CETSA',
  CO10: 'Tesoritos',
  CO12: 'Enerbit',
  Fundación: 'Fundación Celsia',
  'Celsia int': 'Celsia Internet',
  'Pensionados Celsia': 'Pensionados Celsia',
  CO02: 'Cementos Argos',
  CO04: 'Zona Franca',
  Logitrans: 'Logitrans',
  CasaParaMi: 'Casa para mí',
  SolucionesModulares: 'Soluciones Modulares',
  Pensionados: 'Pensionados Cementos',
  CO58: 'Summa',
  OA01: 'Odinsa Aeropuertos',
  OG01: 'Odinsa Gestor',
  OD01: 'Odinsa SA',
  OV01: 'Odinsa Vías',
  '1000': 'Grupo Argos',
  '1600': 'Fundación Grupo Argos',
  '1500': 'Sator',
  '1300': 'Tekia',
}

function titleCaseWords(value: string): string {
  return value
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ')
}

function cleanConceptoLabel(descripcion: string): string {
  let text = descripcion.trim()
  if (!text) return ''
  let changed = true
  while (changed) {
    changed = false
    for (const re of DESCRIPCION_PREFIXES) {
      const next = text.replace(re, '')
      if (next !== text) {
        text = next.trim()
        changed = true
      }
    }
  }
  text = text.replace(/^[.\-\s]+/, '').replace(/[.\-\s]+$/, '').replace(/\s+/g, ' ')
  return titleCaseWords(text)
}

function pickAcreedorLabel(descripciones: string[], acreedorCodigo: string): string {
  const cleaned = descripciones
    .map(cleanConceptoLabel)
    .filter(Boolean)
    .sort((a, b) => a.length - b.length || a.localeCompare(b, 'es'))
  return cleaned[0] || acreedorCodigo
}

function formatPeriodoArchivo(rows: RelacionRow[]): string {
  if (rows.length === 0) return 'periodo'
  const first = rows[0]
  const key = String(first.periodo).trim()
  const q = key === '10' ? '1Q' : key === '20' ? '2Q' : `P${key}`
  const mesFull = MESES_ES[first.mes - 1] ?? String(first.mes)
  const mes = mesFull.slice(0, 3)
  return `${q} ${mes}`
}

function formatEmpresaArchivo(empresa: EmpresaSap): string {
  return EMPRESA_NOMBRE_ARCHIVO[empresa.codigo] ?? titleCaseWords(empresa.nombre)
}

/** Solo caracteres ilegales de Windows; conserva espacios. */
function sanitizeFileName(value: string): string {
  const cleaned = value.trim().replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim()
  return cleaned || 'Relacion'
}

const RELACION_XLSX_BASE_MAX = 60

/** Nombre de pestaña Excel a partir del nombre de archivo (máx. 31; sin \ / ? * [ ]). */
function toExcelSheetName(fileName: string): string {
  const base = fileName.replace(/\.xlsx$/i, '').trim()
  const cleaned = base.replace(/[\\/?*[\]]+/g, ' ').replace(/\s+/g, ' ').trim()
  const sheet = cleaned.slice(0, 31).trim()
  return sheet || 'Relacion'
}

function buildRelacionAcreedorFileName(
  acreedorLabel: string,
  periodoLabel: string,
  empresaLabel: string,
  suffix = '',
): string {
  const raw = `Rel ${acreedorLabel} ${periodoLabel} - ${empresaLabel}${suffix}`
  let base = sanitizeFileName(raw)
  if (base.length > RELACION_XLSX_BASE_MAX) {
    base = base.slice(0, RELACION_XLSX_BASE_MAX).trim()
  }
  return `${base || 'Relacion'}.xlsx`
}

function groupTransferenciaByAcreedor(
  pagoRows: PagoTercerosRow[],
  grupoCatalogo: string,
) {
  const byAcreedor = new Map<
    string,
    { display: string; conceptos: Set<string>; descripciones: string[] }
  >()
  for (const row of pagoRows) {
    if (!includeInRelacionPorAcreedor(row, grupoCatalogo)) continue
    const concepto = row.concepto.trim().toUpperCase()
    if (!concepto) continue

    const acreedorKey = row.acreedor.trim().toUpperCase()
    const isNaAcreedor = !acreedorKey || acreedorKey === 'N/A'
    // Transferencia sin acreedor real no va al ZIP; SOLO SE ENVÍA RELACIÓN sí (por concepto).
    if (isNaAcreedor && !isMetodoSoloSeEnviaRelacion(row.metodoPago)) continue

    const groupKey = isNaAcreedor ? `CONCEPTO:${concepto}` : acreedorKey
    const current = byAcreedor.get(groupKey)
    if (current) {
      current.conceptos.add(concepto)
      if (row.descripcion.trim()) current.descripciones.push(row.descripcion.trim())
    } else {
      byAcreedor.set(groupKey, {
        display: isNaAcreedor
          ? row.descripcion.trim() || row.concepto.trim()
          : row.acreedor.trim(),
        conceptos: new Set([concepto]),
        descripciones: row.descripcion.trim() ? [row.descripcion.trim()] : [],
      })
    }
  }
  return byAcreedor
}

export type TransferenciaRelacionFile = {
  acreedor: string
  fileName: string
  blob: Blob
}

/** Un .xlsx por acreedor (solo PAGO transferencia → filas RELACION de sus conceptos). */
export async function createTransferenciaRelacionFiles(
  relacionRows: RelacionRow[],
  pagoRows: PagoTercerosRow[],
  empresa: EmpresaSap,
  options?: { centralizadorNovedades?: CentralizadorNovedad[] },
): Promise<{ files: TransferenciaRelacionFile[]; filteredRows: RelacionRow[] }> {
  const byAcreedor = groupTransferenciaByAcreedor(pagoRows, empresa.grupoCatalogo)
  const palette = getPaletteForEmpresa(empresa)
  const usedFileNames = new Set<string>()
  const files: TransferenciaRelacionFile[] = []
  const allFiltered: RelacionRow[] = []
  const empresaLabel = formatEmpresaArchivo(empresa)
  const usaGrancoop = empresaUsaGrancoopCentralizador(empresa.codigo)
  const hasGrancoopGroup = [...byAcreedor.values()].some((g) => isGrancoopAcreedor(g.display))
  if (usaGrancoop && hasGrancoopGroup && !options?.centralizadorNovedades) {
    throw new Error(
      'CELSIA COLOMBIA requiere el archivo Centralizador (pestaña Novedades) para la relación GRANCOOP.',
    )
  }

  const acreedores = [...byAcreedor.entries()].sort((a, b) =>
    a[1].display.localeCompare(b[1].display, 'es'),
  )

  for (const [, group] of acreedores) {
    const rows = relacionRows.filter((r) =>
      group.conceptos.has(r.concepto.trim().toUpperCase()),
    )
    if (rows.length === 0) continue

    allFiltered.push(...rows)

    const descripciones =
      group.descripciones.length > 0
        ? group.descripciones
        : rows.map((r) => r.descripcion).filter(Boolean)
    const isGrancoop = usaGrancoop && isGrancoopAcreedor(group.display)
    const acreedorLabel = isGrancoop
      ? 'GRANCOOP'
      : group.display.trim() === '30'
        ? 'Creditos grupo Argos'
        : pickAcreedorLabel(descripciones, group.display)
    const periodoLabel = formatPeriodoArchivo(rows)

    let fileName = buildRelacionAcreedorFileName(acreedorLabel, periodoLabel, empresaLabel)
    let n = 2
    while (usedFileNames.has(fileName.toLowerCase())) {
      fileName = buildRelacionAcreedorFileName(
        acreedorLabel,
        periodoLabel,
        empresaLabel,
        ` (${n})`,
      )
      n += 1
    }
    usedFileNames.add(fileName.toLowerCase())

    const workbook = new ExcelJS.Workbook()
    workbook.creator = 'NOMINA → RELACION POR ACREEDOR'
    workbook.created = new Date()

    if (isGrancoop) {
      const bundle = buildGrancoopBundle(
        rows,
        options!.centralizadorNovedades!,
        empresa.nombre,
      )
      addGrancoopSheets(workbook, bundle, palette)
    } else {
      const periodoTexto = buildPeriodoResumen(rows)
      addRelacionSheet(workbook, rows, empresa, periodoTexto, palette, {
        sheetName: toExcelSheetName(fileName),
        titleText: `RELACIÓN — ${acreedorLabel}`,
        omitMeta: true,
      })

      if (empresaUsaFiaInfConcepto(empresa.codigo) && isFiaAcreedor(group.display)) {
        const fia = buildFiaInfConceptoFromRelacionRows(rows)
        if (fia.rows.length > 0) addFiaInfSheet(workbook, fia, palette)
      }
    }

    files.push({
      acreedor: group.display,
      fileName,
      blob: await workbookToBlob(workbook),
    })
  }

  return { files, filteredRows: allFiltered }
}

/** Carga y parsea Centralizador (Novedades) para GRANCOOP CO05. */
export async function loadCentralizadorNovedades(
  file: File,
): Promise<CentralizadorNovedad[]> {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    throw new Error('El Centralizador debe ser un archivo .xlsx')
  }
  const wb = await loadWorkbookFromFile(file)
  const sheet =
    findSheet(wb, 'Novedades') ??
    wb.worksheets.find((ws) => ws.name.trim().toLowerCase().includes('novedad')) ??
    wb.worksheets[0]
  if (!sheet) {
    throw new Error('No se encontró la pestaña Novedades en el Centralizador.')
  }
  return parseCentralizadorNovedades(worksheetToMatrix(sheet))
}

export async function zipTransferenciaRelacionFiles(
  files: TransferenciaRelacionFile[],
): Promise<Blob> {
  const zip = new JSZip()
  for (const file of files) {
    zip.file(file.fileName, file.blob)
  }
  return zip.generateAsync({ type: 'blob' })
}

export async function loadWorkbookFromFile(file: File): Promise<ExcelJS.Workbook> {
  const buffer = await file.arrayBuffer()
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(buffer)
  return workbook
}

export async function detectPeriodoFromNominaFile(file: File): Promise<number | null> {
  if (!file.name.toLowerCase().endsWith('.xlsx')) return null
  try {
    const source = await loadWorkbookFromFile(file)
    const nomina = findSheet(source, NOMINA_SHEET_NAME)
    if (!nomina) return null
    return detectPeriodoFromNominaMatrix(worksheetToMatrix(nomina))
  } catch {
    return null
  }
}

export async function processNominaFile(
  file: File,
  empresaCodigo: string,
  periodoOverride?: string | number,
  masterFile?: File | null,
): Promise<ProcessOutput> {
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    throw new Error('Solo se admiten archivos .xlsx')
  }

  const empresa = getEmpresaByCodigo(empresaCodigo)
  if (!empresa) {
    throw new Error(`Empresa SAP no reconocida: ${empresaCodigo}`)
  }

  const needsMaster = empresaUsaCentroCostosMaster(empresa.codigo)
  if (needsMaster) {
    if (!masterFile) {
      throw new Error(
        'Fundación Grupo Argos requiere el archivo MASTER (pestaña BD COMPLETA) para CENTRO DE COSTOS.',
      )
    }
    if (!masterFile.name.toLowerCase().endsWith('.xlsx')) {
      throw new Error('El archivo MASTER debe ser .xlsx')
    }
  }

  const source = await loadWorkbookFromFile(file)
  const nomina = findSheet(source, NOMINA_SHEET_NAME)

  if (!nomina) {
    const names = source.worksheets.map((ws) => ws.name).join(', ')
    throw new Error(
      `No se encontró la pestaña "${NOMINA_SHEET_NAME}". Hojas disponibles: ${names || '(ninguna)'}`,
    )
  }

  const matrix = worksheetToMatrix(nomina)

  let costByAssignment: Map<string, ExcelJS.CellValue> | null = null
  let cecoByPersonal: Map<string, string> | undefined
  if (needsMaster && masterFile) {
    try {
      const masterWb = await loadWorkbookFromFile(masterFile)
      costByAssignment = buildAssignmentCostCenterMap(masterWb)
      cecoByPersonal = buildCecoLabelByPersonal(costByAssignment)
    } catch (err) {
      throw new Error(
        err instanceof Error
          ? err.message
          : 'No se pudo leer MASTER para CENTRO DE COSTOS / CECO.',
      )
    }
  }

  const relacion = transformNominaRows(
    matrix,
    empresa.nombre,
    periodoOverride,
    cecoByPersonal,
  )
  const catalogo = getCatalogoByGrupo(empresa.grupoCatalogo)
  const sociedad = codigoEnInforme(empresa)
  const codigoEmpresa = empresa.codigo.trim().toLowerCase()
  const PERIODICIDAD_MENSUAL_FORZADA = new Set(['pensionados', 'pensionados celsia'])
  const PERIODICIDAD_QUINCENAL = new Set([
    'oa01',
    'og01',
    'od01',
    'ov01',
    '1000',
    '1600',
    '1500',
    '1300',
    'co01',
    'co05',
    'cetsa',
    'co10',
    'co12',
    'fundación',
    'co58',
    'co02',
    'co04',
    'logitrans',
    'casaparami',
    'solucionesmodulares',
  ])
  const forcePeriodicidad = PERIODICIDAD_MENSUAL_FORZADA.has(codigoEmpresa)
  const periodicidadDefault = forcePeriodicidad
    ? 'MENSUAL'
    : PERIODICIDAD_QUINCENAL.has(codigoEmpresa)
      ? 'QUINCENAL'
      : ''
  const pago = transformPagoTerceros(
    relacion.rows,
    sociedad,
    catalogo,
    periodicidadDefault,
    forcePeriodicidad,
    empresa.grupoCatalogo,
    needsMaster,
  )
  const pagoConOverrides = applyPeriodicidadOverridesByEmpresa(pago.rows, empresa.codigo)
  // Solo transferencia: suma duplicados exactos acreedor+concepto (antes de CXP)
  let pagoRows = consolidateTransferenciaByAcreedorConcepto(pagoConOverrides)
  let cxpEnrichment: ProcessSummary['cxpEnrichment'] = null
  // Solo pestaña llamada CXP: si no viene, salida con CXP en blanco (no adivinar otras hojas).
  const cxpSheet = findSheet(source, CXP_SHEET_NAME)
  if (cxpSheet) {
    try {
      const enrich = enrichPagoFromCxp(pagoRows, worksheetToMatrix(cxpSheet), sociedad)
      pagoRows = enrich.rows
      cxpEnrichment = {
        fileName: file.name,
        updatedRows: enrich.updatedRows,
        rowsWithDocumento: enrich.rowsWithDocumento,
        cxpRowsUsed: enrich.cxpRowsUsed,
        matchedByAcreedor: enrich.matchedByAcreedor,
        filledWithCompanyDefault: enrich.filledWithCompanyDefault,
        usedSociedadFallback: enrich.usedSociedadFallback,
        missingHeaders: enrich.missingHeaders,
      }
    } catch (err) {
      // No tumbar RELACION/PAGO; avisar en el resumen
      cxpEnrichment = {
        fileName: file.name,
        updatedRows: 0,
        rowsWithDocumento: 0,
        cxpRowsUsed: 0,
        matchedByAcreedor: 0,
        filledWithCompanyDefault: 0,
        error: err instanceof Error ? err.message : 'No se pudo leer la pestaña CXP',
      }
    }
  }

  pagoRows = applyCelsiaInternetPagoDefaults(pagoRows, empresa.codigo)
  pagoRows = applyCementosPagoDefaults(pagoRows, empresa.grupoCatalogo)

  let contingenteSheet: ExcelJS.Worksheet | undefined
  let contingente: ProcessSummary['contingente'] = null
  if (empresaUsaContingenteOrganizado(empresa.grupoCatalogo)) {
    const rawContingente = findSheet(source, CONTINGENTE_SHEET_NAME)
    if (rawContingente) {
      try {
        const variant = contingenteVariantFromGrupo(empresa.grupoCatalogo)
        const organized = organizeContingenteSheet(rawContingente, variant)
        contingenteSheet = organized.sheet
        contingente = { fileName: file.name, dataRows: organized.dataRows }
      } catch (err) {
        contingente = {
          fileName: file.name,
          dataRows: 0,
          error: err instanceof Error ? err.message : 'No se pudo organizar Contingente',
        }
      }
    }
  }

  const workbook = createOutputWorkbook(
    relacion.rows,
    pagoRows,
    empresa,
    nomina,
    cxpSheet,
    contingenteSheet,
  )

  let centroCostos: ProcessSummary['centroCostos'] = null
  if (needsMaster && masterFile && costByAssignment) {
    try {
      const nominaOut = workbook.getWorksheet(NOMINA_SHEET_NAME)
      if (!nominaOut) {
        throw new Error('No se pudo enriquecer NOMINA: hoja ausente en la salida.')
      }
      const stats = enrichNominaWithCentroCostos(nominaOut, costByAssignment)
      centroCostos = {
        fileName: masterFile.name,
        matched: stats.matched,
        totalDataRows: stats.totalDataRows,
      }
    } catch (err) {
      throw new Error(
        err instanceof Error
          ? err.message
          : 'No se pudo aplicar CENTRO DE COSTOS desde MASTER.',
      )
    }
  }

  return {
    workbook,
    relacionRows: relacion.rows,
    pagoRows,
    summary: {
      fileName: file.name,
      empresaCodigo: sociedad,
      empresaNombre: empresa.nombre,
      grupoCatalogo: empresa.grupoCatalogo,
      periodoResumen: buildPeriodoResumen(relacion.rows),
      relacionRows: relacion.detailRows,
      pagoRows: pagoRows.length,
      skippedTotalRows: relacion.skippedTotalRows,
      catalogSize: pago.catalogSize,
      outputSheets: buildOutputSheetNames(true, Boolean(cxpSheet)),
      cxpEnrichment,
      contingente,
      centroCostos,
    },
  }
}

export async function enrichPagoWithCxpFile(params: {
  cxpFile: File
  empresaCodigo: string
  relacionRows: RelacionRow[]
  pagoRows: PagoTercerosRow[]
  baseSummary: ProcessSummary
  nominaSheet?: ExcelJS.Worksheet
  contingenteSheet?: ExcelJS.Worksheet
}): Promise<ProcessOutput & { enrich: EnrichCxpResult }> {
  const {
    cxpFile,
    empresaCodigo,
    relacionRows,
    pagoRows,
    baseSummary,
    nominaSheet,
    contingenteSheet,
  } = params

  if (!cxpFile.name.toLowerCase().endsWith('.xlsx')) {
    throw new Error('Solo se admiten archivos .xlsx para CXP')
  }

  const empresa = getEmpresaByCodigo(empresaCodigo)
  if (!empresa) {
    throw new Error(`Empresa SAP no reconocida: ${empresaCodigo}`)
  }

  const source = await loadWorkbookFromFile(cxpFile)
  const cxp = findCxpSheet(source, { allowFirstSheetFallback: true })
  if (!cxp) {
    const names = source.worksheets.map((ws) => ws.name).join(', ')
    throw new Error(
      `No se encontró pestaña CXP ni hoja con layout Doc/Tercero. Hojas: ${names || '(ninguna)'}`,
    )
  }

  const enrich = enrichPagoFromCxp(pagoRows, worksheetToMatrix(cxp), codigoEnInforme(empresa))
  const pagoFinal = applyCementosPagoDefaults(
    applyCelsiaInternetPagoDefaults(enrich.rows, empresa.codigo),
    empresa.grupoCatalogo,
  )
  const workbook = createOutputWorkbook(
    relacionRows,
    pagoFinal,
    empresa,
    nominaSheet,
    cxp,
    contingenteSheet,
  )

  return {
    workbook,
    relacionRows,
    pagoRows: pagoFinal,
    enrich: { ...enrich, rows: pagoFinal },
    summary: {
      ...baseSummary,
      outputSheets: buildOutputSheetNames(Boolean(nominaSheet), true),
      pagoRows: pagoFinal.length,
      cxpEnrichment: {
        fileName: cxpFile.name,
        updatedRows: enrich.updatedRows,
        rowsWithDocumento: enrich.rowsWithDocumento,
        cxpRowsUsed: enrich.cxpRowsUsed,
        matchedByAcreedor: enrich.matchedByAcreedor,
        filledWithCompanyDefault: enrich.filledWithCompanyDefault,
        usedSociedadFallback: enrich.usedSociedadFallback,
        missingHeaders: enrich.missingHeaders,
      },
    },
  }
}

export async function applyContingenteFile(params: {
  contingenteFile: File
  empresaCodigo: string
  relacionRows: RelacionRow[]
  pagoRows: PagoTercerosRow[]
  baseSummary: ProcessSummary
  nominaSheet?: ExcelJS.Worksheet
  cxpSheet?: ExcelJS.Worksheet
}): Promise<ProcessOutput> {
  const {
    contingenteFile,
    empresaCodigo,
    relacionRows,
    pagoRows,
    baseSummary,
    nominaSheet,
    cxpSheet,
  } = params

  if (!contingenteFile.name.toLowerCase().endsWith('.xlsx')) {
    throw new Error('Solo se admiten archivos .xlsx para Contingente')
  }

  const empresa = getEmpresaByCodigo(empresaCodigo)
  if (!empresa) {
    throw new Error(`Empresa SAP no reconocida: ${empresaCodigo}`)
  }
  if (!empresaUsaContingenteOrganizado(empresa.grupoCatalogo)) {
    throw new Error('Contingente organizado no está disponible para esta empresa.')
  }

  const source = await loadWorkbookFromFile(contingenteFile)
  const raw = findContingenteSheet(source, { allowFirstSheetFallback: true })
  if (!raw) {
    throw new Error('No se encontró hoja Contingente en el archivo.')
  }

  const variant = contingenteVariantFromGrupo(empresa.grupoCatalogo)
  const organized = organizeContingenteSheet(raw, variant)
  const workbook = createOutputWorkbook(
    relacionRows,
    pagoRows,
    empresa,
    nominaSheet,
    cxpSheet,
    organized.sheet,
  )

  return {
    workbook,
    relacionRows,
    pagoRows,
    summary: {
      ...baseSummary,
      outputSheets: buildOutputSheetNames(Boolean(nominaSheet), true),
      contingente: {
        fileName: contingenteFile.name,
        dataRows: organized.dataRows,
      },
    },
  }
}

export async function workbookToBlob(workbook: ExcelJS.Workbook): Promise<Blob> {
  const buffer = await workbook.xlsx.writeBuffer()
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
}

export function downloadBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob)
  const anchor = document.createElement('a')
  anchor.href = url
  anchor.download = fileName
  anchor.click()
  URL.revokeObjectURL(url)
}

type SaveFileHandle = {
  createWritable: () => Promise<{
    write: (data: Blob) => Promise<void>
    close: () => Promise<void>
  }>
}

type SaveFilePickerWindow = Window & {
  showSaveFilePicker?: (options?: {
    suggestedName?: string
    types?: Array<{
      description: string
      accept: Record<string, string[]>
    }>
  }) => Promise<SaveFileHandle>
}

export type SaveTarget =
  | { kind: 'handle'; handle: SaveFileHandle }
  | { kind: 'download'; fileName: string }
  | { kind: 'cancelled' }

function saveFilePickerTypes(fileName: string) {
  if (fileName.toLowerCase().endsWith('.zip')) {
    return [
      {
        description: 'Archivo ZIP',
        accept: { 'application/zip': ['.zip'] },
      },
    ]
  }
  return [
    {
      description: 'Excel',
      accept: {
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      },
    },
  ]
}

/**
 * Abrir "Guardar como" en el gesto del clic (antes de generar el blob).
 * Si no hay API o falla por gesto/permiso → modo descarga automática.
 */
export async function pickSaveTarget(fileName: string): Promise<SaveTarget> {
  const w = window as SaveFilePickerWindow
  if (typeof w.showSaveFilePicker !== 'function') {
    return { kind: 'download', fileName }
  }

  try {
    const handle = await w.showSaveFilePicker({
      suggestedName: fileName,
      types: saveFilePickerTypes(fileName),
    })
    return { kind: 'handle', handle }
  } catch (err) {
    if (err instanceof DOMException && err.name === 'AbortError') {
      return { kind: 'cancelled' }
    }
    if (
      err instanceof DOMException &&
      (err.name === 'NotAllowedError' || err.name === 'SecurityError')
    ) {
      return { kind: 'download', fileName }
    }
    // Otros fallos del picker (p. ej. gesto perdido con mensaje genérico)
    return { kind: 'download', fileName }
  }
}

export async function writeSaveTarget(target: SaveTarget, blob: Blob): Promise<void> {
  if (target.kind === 'cancelled') return
  if (target.kind === 'download') {
    downloadBlob(blob, target.fileName)
    return
  }
  const writable = await target.handle.createWritable()
  await writable.write(blob)
  await writable.close()
}

/** Diálogo Guardar como (Chrome/Edge); fallback a descarga automática. */
export async function saveBlob(blob: Blob, fileName: string): Promise<void> {
  const target = await pickSaveTarget(fileName)
  await writeSaveTarget(target, blob)
}

export function buildOutputFileName(originalName: string, empresaCodigo: string): string {
  const base = originalName.replace(/\.xlsx$/i, '')
  return `TERCEROS_${empresaCodigo}_${base}.xlsx`
}

export function buildTransferenciaZipFileName(
  _originalName: string,
  empresaCodigo: string,
): string {
  return `Relaciones_${empresaCodigo.trim() || 'SAP'}.zip`
}
