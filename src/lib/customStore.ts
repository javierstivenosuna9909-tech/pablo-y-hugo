import type { CatalogoRow, GrupoCatalogo } from '../data/catalogos/types'

const EMPRESAS_KEY = 'terceros.customEmpresas.v1'
const CONCEPTOS_KEY = 'terceros.customConceptos.v1'

export type CustomEmpresa = {
  codigo: string
  nombre: string
  grupoCatalogo: GrupoCatalogo
}

export type CustomConceptosByGrupo = Partial<Record<GrupoCatalogo, CatalogoRow[]>>

function readJson<T>(key: string, fallback: T): T {
  try {
    const raw = localStorage.getItem(key)
    if (!raw) return fallback
    return JSON.parse(raw) as T
  } catch {
    return fallback
  }
}

function writeJson(key: string, value: unknown): void {
  localStorage.setItem(key, JSON.stringify(value))
}

export function loadCustomEmpresas(): CustomEmpresa[] {
  const list = readJson<CustomEmpresa[]>(EMPRESAS_KEY, [])
  return Array.isArray(list) ? list : []
}

export function saveCustomEmpresas(list: CustomEmpresa[]): void {
  writeJson(EMPRESAS_KEY, list)
}

export function addCustomEmpresa(empresa: CustomEmpresa): CustomEmpresa[] {
  const list = loadCustomEmpresas()
  const idx = list.findIndex(
    (e) => e.codigo.trim().toLowerCase() === empresa.codigo.trim().toLowerCase(),
  )
  if (idx >= 0) {
    list[idx] = empresa
  } else {
    list.push(empresa)
  }
  saveCustomEmpresas(list)
  return list
}

export function removeCustomEmpresa(codigo: string): CustomEmpresa[] {
  const next = loadCustomEmpresas().filter(
    (e) => e.codigo.trim().toLowerCase() !== codigo.trim().toLowerCase(),
  )
  saveCustomEmpresas(next)
  return next
}

export function loadCustomConceptos(): CustomConceptosByGrupo {
  const data = readJson<CustomConceptosByGrupo>(CONCEPTOS_KEY, {})
  return data && typeof data === 'object' ? data : {}
}

export function saveCustomConceptos(data: CustomConceptosByGrupo): void {
  writeJson(CONCEPTOS_KEY, data)
}

export function addCustomConcepto(grupo: GrupoCatalogo, row: CatalogoRow): CatalogoRow[] {
  const all = loadCustomConceptos()
  const list = [...(all[grupo] ?? [])]
  const key = row.concepto.trim().toUpperCase()
  const idx = list.findIndex((r) => r.concepto.trim().toUpperCase() === key)
  if (idx >= 0) {
    list[idx] = row
  } else {
    list.push(row)
  }
  all[grupo] = list
  saveCustomConceptos(all)
  return list
}

export function removeCustomConcepto(grupo: GrupoCatalogo, concepto: string): CatalogoRow[] {
  const all = loadCustomConceptos()
  const key = concepto.trim().toUpperCase()
  const list = (all[grupo] ?? []).filter((r) => r.concepto.trim().toUpperCase() !== key)
  all[grupo] = list
  saveCustomConceptos(all)
  return list
}
