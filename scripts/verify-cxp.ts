import ExcelJS from 'exceljs'
import { enrichPagoFromCxp } from '../src/lib/transformCxp'
import { transformPagoTerceros } from '../src/lib/transformPagoTerceros'
import { transformNominaRows } from '../src/lib/transformNomina'
import { getCatalogoByGrupo } from '../src/data/catalogos'

async function main() {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.readFile('NOMINA 2Q OCTUBRE CELSIA SA.xlsx')
  const nomina = wb.worksheets.find((w) => w.name.toLowerCase() === 'nomina')
  const cxp = wb.worksheets.find((w) => w.name.toLowerCase() === 'cxp')
  if (!nomina || !cxp) throw new Error('Faltan hojas')

  const toMatrix = (ws: ExcelJS.Worksheet) => {
    const m: unknown[][] = []
    ws.eachRow({ includeEmpty: false }, (row) => {
      m.push((row.values as unknown[]).slice(1))
    })
    return m
  }

  const relacion = transformNominaRows(toMatrix(nomina), 'CELSIA SA')
  const pago = transformPagoTerceros(
    relacion.rows,
    'CO01',
    getCatalogoByGrupo('CELSIA Y FILIALES'),
  )
  const enrich = enrichPagoFromCxp(pago.rows, toMatrix(cxp), 'CO01')
  const withDoc = enrich.rows.filter(
    (r) => r.numeroDocumento && r.claseDocumento && r.fechaDocumento && r.fechaRecibeCxp,
  )
  const sample = enrich.rows.find((r) => r.concepto === '6MB6') ?? enrich.rows[0]

  console.log(
    JSON.stringify(
      {
        pago: pago.rows.length,
        updated: enrich.updatedRows,
        withDoc: withDoc.length,
        sample,
        fallback: enrich.usedSociedadFallback,
      },
      null,
      2,
    ),
  )
  console.log(withDoc.length === enrich.rows.length ? 'PASS' : 'FAIL')
}

main().catch((err) => {
  console.error(err)
  process.exit(1)
})
