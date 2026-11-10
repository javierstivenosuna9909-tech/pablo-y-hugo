import ExcelJS from 'exceljs'
import fs from 'fs'

const excelPath = 'c:/Users/josuna40/Downloads/CATALOGO X EMPRESA 1.xlsx'
const outDir = 'c:/Users/josuna40/Downloads/Pablo y Hugo 1/src/data/catalogos'

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

function rowsFromSheet(ws) {
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
  return rows
}

function writeCatalog(fileName, exportName, rows) {
  const entries = rows
    .map((r) => '  ' + JSON.stringify(r, null, 2).split('\n').join('\n  '))
    .join(',\n')
  const out = `import type { CatalogoRow } from './types'

export const ${exportName}: CatalogoRow[] = [
${entries}
]
`
  const path = `${outDir}/${fileName}`
  fs.writeFileSync(path, out, 'utf8')
  console.log(`wrote ${rows.length} rows to ${path}`)
}

const wb = new ExcelJS.Workbook()
await wb.xlsx.readFile(excelPath)

const celsia = wb.worksheets.find((w) => w.name.trim().toUpperCase() === 'CELSIA Y FILIALES')
const summaSheet = wb.worksheets.find((w) => w.name.trim().toUpperCase() === 'SUMMA')
if (!celsia || !summaSheet) {
  throw new Error(`Hojas esperadas no encontradas: ${wb.worksheets.map((w) => w.name).join(', ')}`)
}

writeCatalog('celsiaFiliales.ts', 'celsiaFiliales', rowsFromSheet(celsia))
writeCatalog('summa.ts', 'summa', rowsFromSheet(summaSheet))
