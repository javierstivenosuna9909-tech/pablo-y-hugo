import ExcelJS from 'exceljs'
import fs from 'fs'

const excelPath = 'c:/Users/josuna40/Downloads/TERCEROS CELSIA INTERNET 1.xlsx'
const outPath =
  'c:/Users/josuna40/Downloads/Pablo y Hugo 1/src/data/catalogos/celsiaInternet.ts'

function cellText(value) {
  if (value == null) return ''
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim()
  }
  if (value instanceof Date) return value.toISOString()
  if (typeof value === 'object') {
    if (typeof value.text === 'string') return value.text.trim()
    if (value.result != null) return cellText(value.result)
    if (Array.isArray(value.richText)) {
      return value.richText.map((p) => p.text || '').join('').trim()
    }
  }
  return String(value).trim()
}

const wb = new ExcelJS.Workbook()
await wb.xlsx.readFile(excelPath)
const ws = wb.worksheets[0]
const rows = []

for (let r = 2; r <= ws.rowCount; r += 1) {
  const row = ws.getRow(r)
  const acreedor = cellText(row.getCell(1).value)
  const concepto = cellText(row.getCell(2).value)
  const descripcion = cellText(row.getCell(3).value)
  const metodoPago = cellText(row.getCell(4).value)
  const periodicidad = cellText(row.getCell(5).value)
  if (!concepto && !acreedor && !descripcion && !metodoPago) continue
  const item = { acreedor, concepto, descripcion, metodoPago }
  if (periodicidad) item.periodicidad = periodicidad
  rows.push(item)
}

const entries = rows
  .map((r) => '  ' + JSON.stringify(r, null, 2).split('\n').join('\n  '))
  .join(',\n')

const out = `import type { CatalogoRow } from './types'

/** Catálogo Celsia Internet: acreedor = NIT (no código SAP). */
export const celsiaInternet: CatalogoRow[] = [
${entries}
]
`

fs.writeFileSync(outPath, out, 'utf8')
console.log(`wrote ${rows.length} rows to ${outPath}`)
