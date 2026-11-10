import { useMemo, useState } from 'react'
import ExcelJS from 'exceljs'
import { UploadZone } from './components/UploadZone'
import { ResultPanel } from './components/ResultPanel'
import { CatalogAdmin } from './components/CatalogAdmin'
import {
  codigoEnInforme,
  detectEmpresaFromFileName,
  getAllEmpresas,
  getEmpresaByCodigo,
} from './data/empresas'
import {
  applyContingenteFile,
  buildOutputFileName,
  buildTransferenciaZipFileName,
  CONTINGENTE_SHEET_NAME,
  CXP_SHEET_NAME,
  createTransferenciaRelacionFiles,
  detectPeriodoFromNominaFile,
  enrichPagoWithCxpFile,
  loadCentralizadorNovedades,
  NOMINA_SHEET_NAME,
  pickSaveTarget,
  processNominaFile,
  workbookToBlob,
  writeSaveTarget,
  zipTransferenciaRelacionFiles,
  type ProcessSummary,
} from './lib/excelIO'
import { empresaUsaContingenteOrganizado } from './lib/transformContingente'
import { empresaUsaCentroCostosMaster } from './lib/enrichNominaCentroCostos'
import { empresaUsaGrancoopCentralizador } from './lib/transformGrancoop'
import type { RelacionRow } from './lib/transformNomina'
import {
  filterRelacionByTransferenciaConceptos,
  type PagoTercerosRow,
} from './lib/transformPagoTerceros'
import './App.css'

const PERIODO_OPTIONS = Array.from({ length: 24 }, (_, i) => String(i + 1))
const PERIODO_OTRO = 'otro'

function periodoUiFromValue(periodo: string | number): { select: string; otro: string } {
  const text = String(periodo).trim()
  if (PERIODO_OPTIONS.includes(text)) {
    return { select: text, otro: '' }
  }
  return { select: PERIODO_OTRO, otro: text }
}

function parsePeriodoEfectivo(select: string, otro: string): string | number | null {
  if (select === PERIODO_OTRO) {
    const text = otro.trim()
    if (!text) return null
    const asNum = Number(text)
    if (Number.isFinite(asNum) && String(asNum) === text) return asNum
    return text
  }
  if (!PERIODO_OPTIONS.includes(select)) return null
  return Number(select)
}

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [masterFile, setMasterFile] = useState<File | null>(null)
  const [centralizadorFile, setCentralizadorFile] = useState<File | null>(null)
  const [cxpFile, setCxpFile] = useState<File | null>(null)
  const [contingenteFile, setContingenteFile] = useState<File | null>(null)
  const [empresaCodigo, setEmpresaCodigo] = useState('CO01')
  const [periodoSelect, setPeriodoSelect] = useState('10')
  const [periodoOtro, setPeriodoOtro] = useState('')
  const [empresasTick, setEmpresasTick] = useState(0)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [summary, setSummary] = useState<ProcessSummary | null>(null)
  const [workbook, setWorkbook] = useState<ExcelJS.Workbook | null>(null)
  const [relacionRows, setRelacionRows] = useState<RelacionRow[]>([])
  const [pagoRows, setPagoRows] = useState<PagoTercerosRow[]>([])

  const empresas = useMemo(() => {
    void empresasTick
    return getAllEmpresas()
  }, [empresasTick])

  const periodoEfectivo = parsePeriodoEfectivo(periodoSelect, periodoOtro)
  const periodoValido = periodoEfectivo != null
  const empresaActual = useMemo(
    () => getEmpresaByCodigo(empresaCodigo),
    [empresaCodigo, empresasTick],
  )
  const muestraContingente = empresaUsaContingenteOrganizado(
    empresaActual?.grupoCatalogo ?? '',
  )
  const requiereMaster = empresaUsaCentroCostosMaster(empresaCodigo)
  const requiereCentralizador = empresaUsaGrancoopCentralizador(empresaCodigo)

  const resetResult = () => {
    setError(null)
    setSummary(null)
    setWorkbook(null)
    setRelacionRows([])
    setPagoRows([])
    setCxpFile(null)
    setContingenteFile(null)
  }

  const handleFileSelected = async (next: File | null) => {
    setFile(next)
    resetResult()
    if (!next) return
    if (!next.name.toLowerCase().endsWith('.xlsx')) {
      setError('Solo se admiten archivos .xlsx')
      return
    }
    const detected = detectEmpresaFromFileName(next.name)
    if (detected) {
      setEmpresaCodigo(detected.codigo)
      setError(null)
    } else {
      setError(
        'No se pudo detectar la empresa por el nombre del archivo. Selecciónala manualmente en el listado.',
      )
    }

    const detectedPeriodo = await detectPeriodoFromNominaFile(next)
    if (detectedPeriodo != null) {
      const ui = periodoUiFromValue(detectedPeriodo)
      setPeriodoSelect(ui.select)
      setPeriodoOtro(ui.otro)
    }
  }

  const handleGenerate = async () => {
    if (!file) {
      setError('Selecciona un archivo Excel primero.')
      return
    }
    if (periodoEfectivo == null) {
      setError('Indica un periodo válido (1–24 u Otro con texto).')
      return
    }
    if (requiereMaster && !masterFile) {
      setError('Selecciona el archivo MASTER (BD COMPLETA) para Fundación Grupo Argos.')
      return
    }

    setBusy(true)
    setError(null)

    try {
      const output = await processNominaFile(
        file,
        empresaCodigo,
        periodoEfectivo,
        requiereMaster ? masterFile : null,
      )
      setWorkbook(output.workbook)
      setSummary(output.summary)
      setRelacionRows(output.relacionRows)
      setPagoRows(output.pagoRows)
      setCxpFile(null)
      setContingenteFile(null)
    } catch (err) {
      setWorkbook(null)
      setSummary(null)
      setRelacionRows([])
      setPagoRows([])
      setError(err instanceof Error ? err.message : 'No se pudo procesar el archivo.')
    } finally {
      setBusy(false)
    }
  }

  const handleCxpSelected = async (next: File | null) => {
    setCxpFile(next)
    setError(null)

    if (!next) return
    if (!next.name.toLowerCase().endsWith('.xlsx')) {
      setError('Solo se admiten archivos .xlsx para CXP')
      setCxpFile(null)
      return
    }
    if (!summary || pagoRows.length === 0) {
      setError('Primero genera RELACION/PAGO desde NOMINA.')
      return
    }

    setBusy(true)
    try {
      const cont = workbook?.getWorksheet(CONTINGENTE_SHEET_NAME)
      const output = await enrichPagoWithCxpFile({
        cxpFile: next,
        empresaCodigo,
        relacionRows,
        pagoRows,
        baseSummary: summary,
        nominaSheet: workbook?.getWorksheet(NOMINA_SHEET_NAME),
        contingenteSheet: cont && cont.rowCount > 0 ? cont : undefined,
      })
      setWorkbook(output.workbook)
      setPagoRows(output.pagoRows)
      setSummary(output.summary)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'No se pudo aplicar la CXP.')
    } finally {
      setBusy(false)
    }
  }

  const handleContingenteSelected = async (next: File | null) => {
    setContingenteFile(next)
    setError(null)

    if (!next) return
    if (!next.name.toLowerCase().endsWith('.xlsx')) {
      setError('Solo se admiten archivos .xlsx para Contingente')
      setContingenteFile(null)
      return
    }
    if (!summary || !workbook) {
      setError('Primero genera RELACION/PAGO desde NOMINA.')
      return
    }

    setBusy(true)
    try {
      const cxp = workbook.getWorksheet(CXP_SHEET_NAME)
      const output = await applyContingenteFile({
        contingenteFile: next,
        empresaCodigo,
        relacionRows,
        pagoRows,
        baseSummary: summary,
        nominaSheet: workbook.getWorksheet(NOMINA_SHEET_NAME),
        cxpSheet: cxp,
      })
      setWorkbook(output.workbook)
      setSummary(output.summary)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'No se pudo aplicar Contingente.')
    } finally {
      setBusy(false)
    }
  }

  const handleDownload = async () => {
    if (!workbook || !file) return
    const fileName = buildOutputFileName(file.name, empresaCodigo)
    const target = await pickSaveTarget(fileName)
    if (target.kind === 'cancelled') return

    setBusy(true)
    try {
      const blob = await workbookToBlob(workbook)
      await writeSaveTarget(target, blob)
    } catch (err) {
      setError(err instanceof Error ? err.message : 'No se pudo descargar el archivo.')
    } finally {
      setBusy(false)
    }
  }

  const canDownloadTransferencia =
    !busy &&
    Boolean(file) &&
    filterRelacionByTransferenciaConceptos(
      relacionRows,
      pagoRows,
      empresaActual?.grupoCatalogo ?? '',
    ).length > 0

  const handleDownloadTransferencia = async () => {
    if (!file) return
    const empresa = getEmpresaByCodigo(empresaCodigo)
    if (!empresa) {
      setError(`Empresa SAP no reconocida: ${empresaCodigo}`)
      return
    }
    if (requiereCentralizador && !centralizadorFile) {
      setError(
        'Selecciona el archivo Centralizador (Novedades) para generar la relación GRANCOOP.',
      )
      return
    }

    const fileName = buildTransferenciaZipFileName(file.name, empresaCodigo)
    const target = await pickSaveTarget(fileName)
    if (target.kind === 'cancelled') return

    setBusy(true)
    setError(null)
    try {
      const centralizadorNovedades =
        requiereCentralizador && centralizadorFile
          ? await loadCentralizadorNovedades(centralizadorFile)
          : undefined
      const { files, filteredRows } = await createTransferenciaRelacionFiles(
        relacionRows,
        pagoRows,
        empresa,
        { centralizadorNovedades },
      )
      if (filteredRows.length === 0 || files.length === 0) {
        setError(
          'No hay filas de RELACION para acreedores con método transferencia. Nada que unificar.',
        )
        return
      }

      const zipBlob = await zipTransferenciaRelacionFiles(files)
      await writeSaveTarget(target, zipBlob)
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : 'No se pudo descargar las relaciones por acreedor.',
      )
    } finally {
      setBusy(false)
    }
  }

  const hasGenerated = Boolean(summary && workbook)

  return (
    <div className="page">
      <header className="hero">
        <p className="brand">Nómina → Terceros</p>
        <h1>Genera RELACION TERCEROS y PAGO TERCEROS</h1>
        <p className="lede">
          Sube el Excel con NOMINA: la empresa se elige según el nombre del archivo (puedes
          corregirla). Si ya trae pestaña CXP, se completa al generar; si llega después, súbela
          aparte.
        </p>
      </header>

      <main className="stage">
        <section className="controls">
          <UploadZone file={file} disabled={busy} onFileSelected={handleFileSelected} />

          {requiereMaster ? (
            <UploadZone
              file={masterFile}
              disabled={busy}
              onFileSelected={(next) => {
                setMasterFile(next)
                resetResult()
              }}
              title="Arrastra el Excel MASTER"
              hint="Fundación Grupo Argos: pestaña BD COMPLETA (ASSIGNMENT ID → COST CENTER CODE)"
            />
          ) : null}

          {requiereCentralizador ? (
            <UploadZone
              file={centralizadorFile}
              disabled={busy}
              onFileSelected={(next) => {
                setCentralizadorFile(next)
                setError(null)
              }}
              title="Arrastra el Excel Centralizador"
              hint="CELSIA COLOMBIA: pestaña Novedades (requerido para relación GRANCOOP en el ZIP)"
            />
          ) : null}

          <label className="field">
            <span>Empresa</span>
            <select
              value={empresaCodigo}
              onChange={(e) => {
                setEmpresaCodigo(e.target.value)
                resetResult()
              }}
              disabled={busy}
            >
              {empresas.map((empresa) => (
                <option key={empresa.codigo} value={empresa.codigo}>
                  {codigoEnInforme(empresa)} — {empresa.nombre}
                </option>
              ))}
            </select>
          </label>

          <label className="field">
            <span>Periodo</span>
            <select
              value={periodoSelect}
              onChange={(e) => {
                setPeriodoSelect(e.target.value)
                resetResult()
              }}
              disabled={busy}
            >
              {PERIODO_OPTIONS.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
              <option value={PERIODO_OTRO}>Otro…</option>
            </select>
          </label>

          {periodoSelect === PERIODO_OTRO ? (
            <label className="field">
              <span>Periodo extra</span>
              <input
                type="text"
                value={periodoOtro}
                onChange={(e) => {
                  setPeriodoOtro(e.target.value)
                  resetResult()
                }}
                disabled={busy}
                placeholder="Ej. 25 o A1"
              />
            </label>
          ) : null}

          <CatalogAdmin
            disabled={busy}
            selectedEmpresaCodigo={empresaCodigo}
            onEmpresasChange={(selectCodigo) => {
              setEmpresasTick((n) => n + 1)
              const nextList = getAllEmpresas()
              if (selectCodigo) {
                setEmpresaCodigo(selectCodigo)
                resetResult()
              } else if (!nextList.some((e) => e.codigo === empresaCodigo)) {
                setEmpresaCodigo(nextList[0]?.codigo ?? 'CO01')
                resetResult()
              }
            }}
            onCatalogosChange={() => {
              resetResult()
            }}
          />

          <button
            type="button"
            className="btn btn-primary"
            onClick={handleGenerate}
            disabled={busy || !file || !periodoValido || (requiereMaster && !masterFile)}
          >
            {busy && !hasGenerated ? 'Procesando…' : 'Generar pestañas'}
          </button>

          {hasGenerated ? (
            <div className="cxp-block">
              <p className="cxp-label">
                {summary?.cxpEnrichment
                  ? 'CXP ya aplicada — puedes reemplazarla con otro archivo'
                  : 'Paso opcional: CXP aún no vino — súbela cuando llegue'}
              </p>
              <UploadZone
                file={cxpFile}
                disabled={busy}
                onFileSelected={handleCxpSelected}
                title="Arrastra el Excel con CXP"
                hint="Archivo individual con pestaña CXP, o el mismo Excel de nómina que la incluya"
              />
            </div>
          ) : null}

          {hasGenerated && muestraContingente ? (
            <div className="cxp-block">
              <p className="cxp-label">
                {summary?.contingente && !summary.contingente.error
                  ? 'Contingente ya organizado — puedes reemplazarlo'
                  : 'Paso opcional: Contingente — súbelo para organizarlo'}
              </p>
              <UploadZone
                file={contingenteFile}
                disabled={busy}
                onFileSelected={handleContingenteSelected}
                title="Arrastra el Excel Contingente"
                hint="Como lo envían (pestaña CONTINGENTE); se limpian filas/columnas y se colorean aportes"
              />
            </div>
          ) : null}
        </section>

        <ResultPanel
          summary={summary}
          error={error}
          onDownload={handleDownload}
          onDownloadTransferencia={handleDownloadTransferencia}
          canDownload={Boolean(workbook) && !busy}
          canDownloadTransferencia={canDownloadTransferencia}
        />
      </main>
    </div>
  )
}

export default App
