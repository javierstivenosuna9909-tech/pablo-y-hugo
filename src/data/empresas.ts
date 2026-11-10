import type { GrupoCatalogo } from './catalogos/types'
import { loadCustomEmpresas } from '../lib/customStore'

export type EmpresaSap = {
  codigo: string
  nombre: string
  grupoCatalogo: GrupoCatalogo
  /** Código SAP que se escribe en EMPRESA y filtra la sociedad del CXP. Si falta, se usa `codigo`. */
  codigoSociedad?: string
}

export const EMPRESAS_BASE: EmpresaSap[] = [
  { codigo: 'CO01', nombre: 'CELSIA SA', grupoCatalogo: 'CELSIA Y FILIALES' },
  { codigo: 'CO05', nombre: 'CELSIA COLOMBIA', grupoCatalogo: 'CELSIA Y FILIALES' },
  {
    codigo: 'CETSA',
    nombre: 'CETSA',
    grupoCatalogo: 'CELSIA Y FILIALES',
    codigoSociedad: 'CO06',
  },
  { codigo: 'CO10', nombre: 'TESORITOS', grupoCatalogo: 'CELSIA Y FILIALES' },
  { codigo: 'CO12', nombre: 'ENERBIT', grupoCatalogo: 'CELSIA Y FILIALES' },
  { codigo: 'Fundación', nombre: 'FUNDACION CELSIA', grupoCatalogo: 'CELSIA Y FILIALES' },
  { codigo: 'Celsia int', nombre: 'CELSIA INTERNET', grupoCatalogo: 'CELSIA INTERNET' },
  {
    codigo: 'Pensionados Celsia',
    nombre: 'PENSIONADOS CELSIA',
    grupoCatalogo: 'CELSIA Y FILIALES',
    codigoSociedad: 'CO05',
  },
  { codigo: 'CO02', nombre: 'CEMENTOS', grupoCatalogo: 'CEMENTOS Y FILIALES' },
  { codigo: 'CO04', nombre: 'ZONA FRANCA', grupoCatalogo: 'CEMENTOS Y FILIALES' },
  {
    codigo: 'Logitrans',
    nombre: 'LOGITRANS',
    grupoCatalogo: 'CEMENTOS Y FILIALES',
    codigoSociedad: 'CO06',
  },
  {
    codigo: 'CasaParaMi',
    nombre: 'CASA PARA MI',
    grupoCatalogo: 'CEMENTOS Y FILIALES',
  },
  {
    codigo: 'SolucionesModulares',
    nombre: 'SOLUCIONES MODULARES',
    grupoCatalogo: 'CEMENTOS Y FILIALES',
  },
  {
    codigo: 'Pensionados',
    nombre: 'PENSIONADOS CEMENTOS',
    grupoCatalogo: 'CEMENTOS Y FILIALES',
  },
  { codigo: 'CO58', nombre: 'SUMMA', grupoCatalogo: 'SUMMA' },
  { codigo: 'OA01', nombre: 'ODINSA AEROPUERTOS', grupoCatalogo: 'ODINSA Y FILIALES' },
  { codigo: 'OG01', nombre: 'ODINSA GESTOR', grupoCatalogo: 'ODINSA Y FILIALES' },
  { codigo: 'OD01', nombre: 'ODINSA SA', grupoCatalogo: 'ODINSA Y FILIALES' },
  { codigo: 'OV01', nombre: 'ODINSA VIAS', grupoCatalogo: 'ODINSA Y FILIALES' },
  { codigo: '1000', nombre: 'GRUPO ARGOS', grupoCatalogo: 'GRUPO ARGOS Y FILIALES' },
  { codigo: '1600', nombre: 'FUNDACION GRUPO ARGOS', grupoCatalogo: 'GRUPO ARGOS Y FILIALES' },
  { codigo: '1500', nombre: 'SATOR', grupoCatalogo: 'GRUPO ARGOS Y FILIALES' },
  { codigo: '1300', nombre: 'TEKIA', grupoCatalogo: 'GRUPO ARGOS Y FILIALES' },
]

/** @deprecated use getAllEmpresas() — kept for compatibility */
export const EMPRESAS = EMPRESAS_BASE

export function getAllEmpresas(): EmpresaSap[] {
  const map = new Map<string, EmpresaSap>()
  for (const empresa of EMPRESAS_BASE) {
    map.set(empresa.codigo.trim().toLowerCase(), empresa)
  }
  for (const empresa of loadCustomEmpresas()) {
    map.set(empresa.codigo.trim().toLowerCase(), empresa)
  }
  return [...map.values()]
}

export function getEmpresaByCodigo(codigo: string): EmpresaSap | undefined {
  const target = codigo.trim().toLowerCase()
  return getAllEmpresas().find((e) => e.codigo.trim().toLowerCase() === target)
}

export function isBaseEmpresa(codigo: string): boolean {
  const target = codigo.trim().toLowerCase()
  return EMPRESAS_BASE.some((e) => e.codigo.trim().toLowerCase() === target)
}

/** Código que va en la columna EMPRESA y en el filtro de sociedad del CXP. */
export function codigoEnInforme(empresa: EmpresaSap): string {
  const sociedad = empresa.codigoSociedad?.trim()
  return sociedad || empresa.codigo
}

function normalizeText(value: string): string {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/** Alias adicionales frecuentes en nombres de archivo de nómina. */
const FILE_NAME_ALIASES: Array<{ codigo: string; aliases: string[] }> = [
  { codigo: 'CO01', aliases: ['celsia sa', 'celsia s a', 'celsia s.a'] },
  { codigo: 'CO05', aliases: ['celsia colombia'] },
  { codigo: 'CETSA', aliases: ['cetsa'] },
  { codigo: 'Logitrans', aliases: ['logitrans'] },
  { codigo: 'CO10', aliases: ['tesoritos'] },
  { codigo: 'CO12', aliases: ['enerbit'] },
  { codigo: 'Fundación', aliases: ['fundacion celsia', 'fundacion'] },
  { codigo: 'Celsia int', aliases: ['celsia internet', 'celsia int'] },
  { codigo: 'Pensionados Celsia', aliases: ['pensionados celsia'] },
  { codigo: 'CO02', aliases: ['cementos'] },
  { codigo: 'CO04', aliases: ['zona franca'] },
  { codigo: 'CasaParaMi', aliases: ['casa para mi'] },
  { codigo: 'SolucionesModulares', aliases: ['soluciones modulares'] },
  { codigo: 'Pensionados', aliases: ['pensionados cementos', 'pensionados'] },
  { codigo: 'CO58', aliases: ['summa'] },
  { codigo: 'OA01', aliases: ['odinsa aeropuertos'] },
  { codigo: 'OG01', aliases: ['odinsa gestor'] },
  { codigo: 'OD01', aliases: ['odinsa sa', 'odinsa s a'] },
  { codigo: 'OV01', aliases: ['odinsa vias'] },
  { codigo: '1000', aliases: ['grupo argos'] },
  { codigo: '1600', aliases: ['fundacion grupo argos'] },
  { codigo: '1500', aliases: ['sator'] },
  { codigo: '1300', aliases: ['tekia'] },
]

/**
 * Detecta la empresa a partir del nombre del archivo (.xlsx).
 * Prioriza: código SAP/interno completo → alias/nombre más largo.
 */
export function detectEmpresaFromFileName(fileName: string): EmpresaSap | undefined {
  const baseName = normalizeText(fileName.replace(/\.xlsx$/i, ''))
  if (!baseName) return undefined

  const empresas = getAllEmpresas()
  const byCodigo = new Map(empresas.map((e) => [e.codigo.trim().toLowerCase(), e]))

  // CO05 también es CELSIA COLOMBIA; "pensionados celsia" es más específico.
  if (baseName.includes('pensionados celsia')) {
    const pensionadosCelsia = byCodigo.get('pensionados celsia')
    if (pensionadosCelsia) return pensionadosCelsia
  }

  // 1) Código exacto como token (CO01, 1000, fundacion, etc.)
  for (const empresa of empresas) {
    const code = normalizeText(empresa.codigo)
    if (!code) continue
    const re = new RegExp(`(?:^|\\s)${escapeRegExp(code)}(?:\\s|$)`)
    if (re.test(baseName)) return empresa
  }

  // 2) Alias + nombre, del más específico (largo) al más corto
  const candidates: Array<{ alias: string; empresa: EmpresaSap }> = []
  for (const empresa of empresas) {
    candidates.push({ alias: normalizeText(empresa.nombre), empresa })
    candidates.push({ alias: normalizeText(empresa.codigo), empresa })
  }
  for (const entry of FILE_NAME_ALIASES) {
    const empresa = byCodigo.get(entry.codigo.trim().toLowerCase())
    if (!empresa) continue
    for (const alias of entry.aliases) {
      candidates.push({ alias: normalizeText(alias), empresa })
    }
  }

  candidates.sort((a, b) => b.alias.length - a.alias.length)

  for (const { alias, empresa } of candidates) {
    if (!alias) continue
    if (baseName.includes(alias)) return empresa
  }

  // 3) Fallback genérico "celsia" → CELSIA SA (si no matcheó una filial más específica)
  if (/\bcelsia\b/.test(baseName)) {
    return byCodigo.get('co01')
  }

  return undefined
}
