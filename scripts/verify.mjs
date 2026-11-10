import ExcelJS from 'exceljs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const root = path.resolve(__dirname, '..')
const sample = path.join(root, 'NOMINA 2Q OCTUBRE CELSIA SA.xlsx')

function normalizeHeader(value) {
  return String(value ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ')
}

function findColumnIndex(headers, aliases) {
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

function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569)
  return new Date(utcDays * 86400 * 1000)
}

function cellToDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value
  if (typeof value === 'number' && Number.isFinite(value)) return excelSerialToDate(value)
  return null
}

function worksheetToMatrix(worksheet) {
  const matrix = []
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    matrix.push(row.values.slice(1))
  })
  return matrix
}

const workbook = new ExcelJS.Workbook()
await workbook.xlsx.readFile(sample)

const nomina = workbook.worksheets.find((ws) => ws.name.trim().toLowerCase() === 'nomina')
const existingRel = workbook.worksheets.find(
  (ws) => ws.name.trim().toLowerCase() === 'relacion terceros',
)

if (!nomina) throw new Error('NOMINA missing')

const matrix = worksheetToMatrix(nomina)
const headers = matrix[0].map((h) => String(h ?? ''))
const indexes = {
  nombre: findColumnIndex(headers, ['nombre empl./cand.', 'nombre']),
  cedula: findColumnIndex(headers, ['número id', 'numero id']),
  concepto: findColumnIndex(headers, ['cc-nómina', 'cc-nomina']),
  descripcion: findColumnIndex(headers, ['texto expl.cc-nómina', 'texto expl.cc-nomina']),
  neto: findColumnIndex(headers, ['total']),
  fechaFin: findColumnIndex(headers, ['per.para, fecha fin']),
}

console.log('indexes', indexes)

const generated = []
let skipped = 0
for (let i = 1; i < matrix.length; i += 1) {
  const row = matrix[i] ?? []
  const nombre = row[indexes.nombre]
  if (nombre == null || String(nombre).trim() === '') {
    skipped += 1
    continue
  }
  const fecha = cellToDate(row[indexes.fechaFin])
  generated.push({
    anio: fecha.getUTCFullYear(),
    mes: fecha.getUTCMonth() + 1,
    periodo: fecha.getUTCDate() <= 15 ? 10 : 20,
    cedula: String(row[indexes.cedula]),
    nombre: String(nombre).trim(),
    concepto: String(row[indexes.concepto] ?? '').trim(),
    descripcion: String(row[indexes.descripcion] ?? '').trim(),
    neto: Number(row[indexes.neto]),
  })
}

const expected = []
existingRel.eachRow({ includeEmpty: false }, (row, rowNumber) => {
  if (rowNumber === 1) return
  const v = row.values
  expected.push({
    anio: Number(v[1]),
    mes: Number(v[2]),
    periodo: Number(v[3]),
    cedula: String(v[4]),
    nombre: String(v[5]).trim(),
    concepto: String(v[6]).trim(),
    descripcion: String(v[7]).trim(),
    neto: Number(v[8]),
  })
})

console.log('generated', generated.length, 'expected', expected.length, 'skipped', skipped)

let mismatches = 0
for (let i = 0; i < Math.min(generated.length, expected.length); i += 1) {
  const a = generated[i]
  const b = expected[i]
  if (
    a.anio !== b.anio ||
    a.mes !== b.mes ||
    a.periodo !== b.periodo ||
    a.cedula !== b.cedula ||
    a.nombre !== b.nombre ||
    a.concepto !== b.concepto ||
    a.descripcion !== b.descripcion ||
    a.neto !== b.neto
  ) {
    mismatches += 1
    if (mismatches <= 5) console.log('mismatch', i + 1, a, b)
  }
}

console.log('mismatches', mismatches)

// write output with replaced sheet
if (existingRel) workbook.removeWorksheet(existingRel.id)
const sheet = workbook.addWorksheet('RELACION TERCEROS')
sheet.addRow(['AÑO', 'MES', 'PERIODO', 'CEDULA', 'NOMBRE_COMPLETO', 'CONCEPTO', 'DESCRIPCION', 'NETO', 'SOCIEDAD'])
for (const row of generated) {
  sheet.addRow([
    row.anio,
    row.mes,
    row.periodo,
    row.cedula,
    row.nombre,
    row.concepto,
    row.descripcion,
    row.neto,
    'CELSIA S.A.',
  ])
}

const out = path.join(root, 'scripts', '_verify_out.xlsx')
await workbook.xlsx.writeFile(out)
console.log('wrote', out)
console.log('sheets', workbook.worksheets.map((w) => w.name))
