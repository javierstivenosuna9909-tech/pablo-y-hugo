import type { GrupoCatalogo } from './catalogos/types'
import type { EmpresaSap } from './empresas'

/** Hex sin # — listos para ExcelJS argb */
export type BrandPalette = {
  titleBg: string
  titleFg: string
  metaBg: string
  metaFg: string
  metaAccent: string
  headerBg: string
  headerFg: string
  zebra: string
  border: string
  ink: string
}

/** House of Brands — Brandcenter Argos (Cementos) */
export const ARGOS_BLUE = '071D49'
export const ARGOS_GREEN = 'C4D600'
export const ARGOS_WHITE = 'FFFFFF'

/** Celsia — naranja de marca */
export const CELSIA_ORANGE = 'FF812D'
export const CELSIA_ROPE = '995522'
export const CELSIA_INK = '2B1500'

/** Odinsa — Cool Gray 11 + Pantone 633C */
export const ODINSA_GRAY = '555559'
export const ODINSA_TEAL = '007197'

/** Summa — verde de marca #2F4F3E */
export const SUMMA_GREEN = '2F4F3E'

const STANDARD: BrandPalette = {
  titleBg: ARGOS_BLUE,
  titleFg: ARGOS_WHITE,
  metaBg: 'E8ECF2',
  metaFg: ARGOS_BLUE,
  metaAccent: ARGOS_GREEN,
  headerBg: ARGOS_BLUE,
  headerFg: ARGOS_WHITE,
  zebra: 'F3F5F8',
  border: 'C5CCD8',
  ink: ARGOS_BLUE,
}

/** Celsia / energía: naranja oficial #FF812D */
const CELSIA: BrandPalette = {
  titleBg: CELSIA_ORANGE,
  titleFg: CELSIA_INK,
  metaBg: 'FFF3E8',
  metaFg: CELSIA_ROPE,
  metaAccent: CELSIA_ORANGE,
  headerBg: CELSIA_ORANGE,
  headerFg: CELSIA_INK,
  zebra: 'FFF8F0',
  border: 'F0C9A8',
  ink: CELSIA_ROPE,
}

/** Odinsa: Cool Gray + acento teal (sin verde Argos) */
const ODINSA: BrandPalette = {
  titleBg: ODINSA_GRAY,
  titleFg: ARGOS_WHITE,
  metaBg: 'EEF0F2',
  metaFg: ODINSA_GRAY,
  metaAccent: ODINSA_TEAL,
  headerBg: ODINSA_GRAY,
  headerFg: ARGOS_WHITE,
  zebra: 'F5F5F5',
  border: 'C8CDD1',
  ink: ODINSA_GRAY,
}

/** Summa: verde #2F4F3E */
const SUMMA: BrandPalette = {
  titleBg: SUMMA_GREEN,
  titleFg: ARGOS_WHITE,
  metaBg: 'E8EFEA',
  metaFg: SUMMA_GREEN,
  metaAccent: SUMMA_GREEN,
  headerBg: SUMMA_GREEN,
  headerFg: ARGOS_WHITE,
  zebra: 'F2F6F3',
  border: 'C5D0C9',
  ink: SUMMA_GREEN,
}

const BY_CODIGO: Record<string, BrandPalette> = {
  // Grupo Argos / fundación
  '1000': STANDARD,
  '1600': STANDARD,
  // Cementos / zona franca / apoyo
  co02: STANDARD,
  co04: STANDARD,
  logitrans: STANDARD,
  co58: SUMMA,
  '1500': STANDARD,
  '1300': STANDARD,
  co10: STANDARD,
  // Celsia y filiales energía
  co01: CELSIA,
  co05: CELSIA,
  cetsa: CELSIA,
  co12: CELSIA,
  fundación: CELSIA,
  'celsia int': CELSIA,
  // Odinsa
  od01: ODINSA,
  oa01: ODINSA,
  og01: ODINSA,
  ov01: ODINSA,
}

const BY_GRUPO: Record<GrupoCatalogo, BrandPalette> = {
  'CELSIA Y FILIALES': CELSIA,
  'CELSIA INTERNET': CELSIA,
  'CEMENTOS Y FILIALES': STANDARD,
  SUMMA: SUMMA,
  'GRUPO ARGOS Y FILIALES': STANDARD,
  'ODINSA Y FILIALES': ODINSA,
}

export function getPaletteForEmpresa(empresa: EmpresaSap): BrandPalette {
  const byCode = BY_CODIGO[empresa.codigo.trim().toLowerCase()]
  if (byCode) return byCode
  return BY_GRUPO[empresa.grupoCatalogo] ?? STANDARD
}
