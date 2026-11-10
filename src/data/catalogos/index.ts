import type { CatalogoRow, GrupoCatalogo } from './types'
import { celsiaFiliales } from './celsiaFiliales'
import { celsiaInternet } from './celsiaInternet'
import { odinsaFiliales } from './odinsaFiliales'
import { summa } from './summa'
import { grupoArgosFiliales } from './grupoArgosFiliales'
import { cementosFiliales } from './cementosFiliales'
import { loadCustomConceptos } from '../../lib/customStore'

const CATALOGOS: Record<GrupoCatalogo, CatalogoRow[]> = {
  'CELSIA Y FILIALES': celsiaFiliales,
  'CELSIA INTERNET': celsiaInternet,
  'ODINSA Y FILIALES': odinsaFiliales,
  SUMMA: summa,
  'GRUPO ARGOS Y FILIALES': grupoArgosFiliales,
  'CEMENTOS Y FILIALES': cementosFiliales,
}

export function getCatalogoByGrupo(grupo: GrupoCatalogo): CatalogoRow[] {
  const map = new Map<string, CatalogoRow>()
  for (const row of CATALOGOS[grupo] ?? []) {
    map.set(row.concepto.trim().toUpperCase(), row)
  }
  for (const row of loadCustomConceptos()[grupo] ?? []) {
    // Custom overrides / adds
    map.set(row.concepto.trim().toUpperCase(), row)
  }
  return [...map.values()]
}

export function catalogoByConcepto(grupo: GrupoCatalogo): Map<string, CatalogoRow> {
  const map = new Map<string, CatalogoRow>()
  for (const row of getCatalogoByGrupo(grupo)) {
    const key = row.concepto.trim().toUpperCase()
    if (!map.has(key)) {
      map.set(key, row)
    }
  }
  return map
}

export function getCustomConceptosByGrupo(grupo: GrupoCatalogo): CatalogoRow[] {
  return loadCustomConceptos()[grupo] ?? []
}

export type { CatalogoRow, GrupoCatalogo }
export { GRUPOS_CATALOGO } from './types'
