import { enrichPagoFromCxp } from '../src/lib/transformCxp'
import type { PagoTercerosRow } from '../src/lib/transformPagoTerceros'

const matrix = [
  ['Sociedad', 'Proveedor', 'Clase de documento', 'Fecha de documento'],
  ['CO01', '104744', 'AB', 46322],
]

const pago: PagoTercerosRow[] = [
  {
    empresa: 'CO01',
    acreedor: '104744',
    concepto: '6MB6',
    descripcion: 'X',
    valorEnNomina: -1,
    numeroDocumento: '',
    claseDocumento: '',
    fechaDocumento: '',
    metodoPago: 'PLANO',
    valorCxp: '',
    diferencia: '',
    fechaRecibeCxp: '',
    fechaEnvioPago: '',
    seEnviaRelacion: '',
    periodicidad: '',
  },
]

const enrich = enrichPagoFromCxp(pago, matrix, 'CO01')
const row = enrich.rows[0]
console.log(row)
console.log('missing', enrich.missingHeaders)
const ok =
  row.numeroDocumento === '' &&
  row.claseDocumento === 'AB' &&
  Boolean(row.fechaDocumento) &&
  row.fechaRecibeCxp === row.fechaDocumento
console.log(ok ? 'PASS partial' : 'FAIL')
