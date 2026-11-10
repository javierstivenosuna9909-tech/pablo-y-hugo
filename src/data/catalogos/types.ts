export type CatalogoRow = {
  acreedor: string
  concepto: string
  descripcion: string
  metodoPago: string
  /** Si existe, se usa en PAGO TERCEROS (salvo empresas que fuerzan periodicidad). */
  periodicidad?: string
}

export const GRUPOS_CATALOGO = [
  'CELSIA Y FILIALES',
  'CELSIA INTERNET',
  'ODINSA Y FILIALES',
  'SUMMA',
  'GRUPO ARGOS Y FILIALES',
  'CEMENTOS Y FILIALES',
] as const

export type GrupoCatalogo = (typeof GRUPOS_CATALOGO)[number]
