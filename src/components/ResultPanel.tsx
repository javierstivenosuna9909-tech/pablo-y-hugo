import type { ProcessSummary } from '../lib/excelIO'

type ResultPanelProps = {
  summary: ProcessSummary | null
  error: string | null
  onDownload: () => void
  onDownloadTransferencia: () => void
  canDownload: boolean
  canDownloadTransferencia: boolean
}

export function ResultPanel({
  summary,
  error,
  onDownload,
  onDownloadTransferencia,
  canDownload,
  canDownloadTransferencia,
}: ResultPanelProps) {
  if (error) {
    return (
      <div className="panel panel-error" role="alert">
        <h2>Error</h2>
        <p>{error}</p>
      </div>
    )
  }

  if (!summary) {
    return (
      <div className="panel panel-muted">
        <h2>Resultado</h2>
        <p>Sube un archivo, elige la empresa y genera para ver el resumen aquí.</p>
      </div>
    )
  }

  const cxp = summary.cxpEnrichment
  const contingente = summary.contingente
  const centroCostos = summary.centroCostos

  return (
    <div className="panel panel-success">
      <h2>Listo</h2>
      <ul className="summary-list">
        <li>
          <span>Origen</span>
          <strong>{summary.fileName}</strong>
        </li>
        <li>
          <span>Empresa</span>
          <strong>
            {summary.empresaCodigo} — {summary.empresaNombre}
          </strong>
        </li>
        <li>
          <span>Catálogo</span>
          <strong>
            {summary.grupoCatalogo} ({summary.catalogSize} conceptos)
          </strong>
        </li>
        <li>
          <span>Periodo</span>
          <strong>{summary.periodoResumen}</strong>
        </li>
        <li>
          <span>RELACION TERCEROS</span>
          <strong>{summary.relacionRows} filas</strong>
        </li>
        <li>
          <span>PAGO TERCEROS</span>
          <strong>{summary.pagoRows} filas</strong>
        </li>
        <li>
          <span>Totales omitidos en nómina</span>
          <strong>{summary.skippedTotalRows}</strong>
        </li>
        <li>
          <span>Pestañas de salida</span>
          <strong>{(summary.outputSheets ?? []).join(', ')}</strong>
        </li>
        {!cxp ? (
          <li>
            <span>CXP</span>
            <strong>Pendiente — súbela aparte o en un Excel con pestaña CXP</strong>
          </li>
        ) : cxp.error ? (
          <li>
            <span>CXP</span>
            <strong className="warn-text">Error: {cxp.error}</strong>
          </li>
        ) : (
          <>
            <li>
              <span>CXP cargada</span>
              <strong>{cxp.fileName}</strong>
            </li>
            <li>
              <span>Filas PAGO con Nº documento</span>
              <strong>
                {cxp.rowsWithDocumento} / {summary.pagoRows}
              </strong>
            </li>
            <li>
              <span>Filas PAGO actualizadas</span>
              <strong>{cxp.updatedRows}</strong>
            </li>
            <li>
              <span>Match por acreedor</span>
              <strong>{cxp.matchedByAcreedor}</strong>
            </li>
            <li>
              <span>Completadas con default CXP</span>
              <strong>{cxp.filledWithCompanyDefault}</strong>
            </li>
            {cxp.usedSociedadFallback ? (
              <li>
                <span>Aviso CXP</span>
                <strong className="warn-text">
                  No hubo filas con Sociedad={summary.empresaCodigo}; se usó todo el CXP
                </strong>
              </li>
            ) : null}
            {cxp.missingHeaders && cxp.missingHeaders.length > 0 ? (
              <li>
                <span>Encabezados CXP no encontrados</span>
                <strong className="warn-text">
                  {cxp.missingHeaders.join(', ')} (quedan en blanco)
                </strong>
              </li>
            ) : null}
          </>
        )}
        {!contingente ? (
          <li>
            <span>Contingente</span>
            <strong>Pendiente — súbelo aparte para organizarlo</strong>
          </li>
        ) : contingente.error ? (
          <li>
            <span>Contingente</span>
            <strong className="warn-text">Error: {contingente.error}</strong>
          </li>
        ) : (
          <>
            <li>
              <span>Contingente</span>
              <strong>{contingente.fileName}</strong>
            </li>
            <li>
              <span>Filas Contingente</span>
              <strong>{contingente.dataRows}</strong>
            </li>
          </>
        )}
        {centroCostos ? (
          centroCostos.error ? (
            <li>
              <span>CENTRO DE COSTOS</span>
              <strong className="warn-text">Error: {centroCostos.error}</strong>
            </li>
          ) : (
            <>
              <li>
                <span>MASTER (centro costos)</span>
                <strong>{centroCostos.fileName}</strong>
              </li>
              <li>
                <span>Matches centro costos</span>
                <strong>
                  {centroCostos.matched} / {centroCostos.totalDataRows}
                </strong>
              </li>
            </>
          )
        ) : null}
      </ul>
      <div className="download-actions">
        <button type="button" className="btn btn-primary" onClick={onDownload} disabled={!canDownload}>
          Descargar Excel
        </button>
        <button
          type="button"
          className="btn btn-secondary"
          onClick={onDownloadTransferencia}
          disabled={!canDownloadTransferencia}
        >
          Descargar relaciones por acreedor
        </button>
      </div>
    </div>
  )
}
