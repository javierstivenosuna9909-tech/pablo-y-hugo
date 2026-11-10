import { useState, type FormEvent } from 'react'
import { GRUPOS_CATALOGO, type CatalogoRow, type GrupoCatalogo } from '../data/catalogos'
import { isBaseEmpresa, type EmpresaSap } from '../data/empresas'
import {
  addCustomConcepto,
  addCustomEmpresa,
  loadCustomConceptos,
  loadCustomEmpresas,
  removeCustomConcepto,
  removeCustomEmpresa,
} from '../lib/customStore'

type CatalogAdminProps = {
  disabled?: boolean
  selectedEmpresaCodigo: string
  onEmpresasChange: (selectCodigo?: string) => void
  onCatalogosChange: () => void
}

export function CatalogAdmin({
  disabled,
  selectedEmpresaCodigo,
  onEmpresasChange,
  onCatalogosChange,
}: CatalogAdminProps) {
  const [open, setOpen] = useState(false)
  const [message, setMessage] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [tick, setTick] = useState(0)

  const [empCodigo, setEmpCodigo] = useState('')
  const [empNombre, setEmpNombre] = useState('')
  const [empGrupo, setEmpGrupo] = useState<GrupoCatalogo>('CELSIA Y FILIALES')

  const [conceptoGrupo, setConceptoGrupo] = useState<GrupoCatalogo>('CELSIA Y FILIALES')
  const [acreedor, setAcreedor] = useState('')
  const [concepto, setConcepto] = useState('')
  const [descripcion, setDescripcion] = useState('')
  const [metodoPago, setMetodoPago] = useState('')

  // tick forces re-read from localStorage after mutations
  void tick
  const customEmpresas = loadCustomEmpresas()
  const customConceptos = loadCustomConceptos()[conceptoGrupo] ?? []

  const refresh = () => setTick((n) => n + 1)

  const flash = (ok: string | null, err: string | null) => {
    setMessage(ok)
    setError(err)
  }

  const handleAddEmpresa = (event: FormEvent) => {
    event.preventDefault()
    const codigo = empCodigo.trim()
    const nombre = empNombre.trim()
    if (!codigo || !nombre) {
      flash(null, 'Código y nombre de empresa son obligatorios.')
      return
    }
    if (isBaseEmpresa(codigo)) {
      flash(null, 'Ese código ya existe en el listado base; usa otro código interno.')
      return
    }

    const empresa: EmpresaSap = {
      codigo,
      nombre: nombre.toUpperCase(),
      grupoCatalogo: empGrupo,
    }
    addCustomEmpresa(empresa)
    refresh()
    onEmpresasChange(codigo)
    setEmpCodigo('')
    setEmpNombre('')
    flash(`Empresa "${codigo}" agregada.`, null)
  }

  const handleRemoveEmpresa = (codigo: string) => {
    removeCustomEmpresa(codigo)
    refresh()
    onEmpresasChange(selectedEmpresaCodigo === codigo ? undefined : selectedEmpresaCodigo)
    flash(`Empresa "${codigo}" eliminada.`, null)
  }

  const handleAddConcepto = (event: FormEvent) => {
    event.preventDefault()
    const row: CatalogoRow = {
      acreedor: acreedor.trim(),
      concepto: concepto.trim(),
      descripcion: descripcion.trim(),
      metodoPago: metodoPago.trim(),
    }
    if (!row.concepto) {
      flash(null, 'El concepto es obligatorio.')
      return
    }
    if (!row.acreedor || !row.metodoPago) {
      flash(null, 'Acreedor y método de pago son obligatorios.')
      return
    }

    addCustomConcepto(conceptoGrupo, row)
    refresh()
    onCatalogosChange()
    setAcreedor('')
    setConcepto('')
    setDescripcion('')
    setMetodoPago('')
    flash(`Concepto "${row.concepto}" agregado en ${conceptoGrupo}.`, null)
  }

  const handleRemoveConcepto = (conceptoCode: string) => {
    removeCustomConcepto(conceptoGrupo, conceptoCode)
    refresh()
    onCatalogosChange()
    flash(`Concepto "${conceptoCode}" eliminado.`, null)
  }

  return (
    <div className="admin-block">
      <button
        type="button"
        className="btn btn-secondary"
        disabled={disabled}
        onClick={() => setOpen((v) => !v)}
      >
        {open ? 'Ocultar catálogo' : 'Agregar empresas / conceptos'}
      </button>

      {open ? (
        <div className="admin-panel">
          {message ? <p className="admin-ok">{message}</p> : null}
          {error ? <p className="admin-error">{error}</p> : null}

          <form className="admin-form" onSubmit={handleAddEmpresa}>
            <h3>Nueva empresa</h3>
            <div className="admin-grid">
              <label className="field">
                <span>Código</span>
                <input
                  value={empCodigo}
                  onChange={(e) => setEmpCodigo(e.target.value)}
                  disabled={disabled}
                  placeholder="Ej. Fundación"
                />
              </label>
              <label className="field">
                <span>Nombre</span>
                <input
                  value={empNombre}
                  onChange={(e) => setEmpNombre(e.target.value)}
                  disabled={disabled}
                  placeholder="Ej. FUNDACION CELSIA"
                />
              </label>
              <label className="field">
                <span>Catálogo / grupo</span>
                <select
                  value={empGrupo}
                  onChange={(e) => setEmpGrupo(e.target.value as GrupoCatalogo)}
                  disabled={disabled}
                >
                  {GRUPOS_CATALOGO.map((grupo) => (
                    <option key={grupo} value={grupo}>
                      {grupo}
                    </option>
                  ))}
                </select>
              </label>
            </div>
            <button type="submit" className="btn btn-primary" disabled={disabled}>
              Guardar empresa
            </button>
          </form>

          {customEmpresas.length > 0 ? (
            <ul className="admin-list">
              {customEmpresas.map((empresa) => (
                <li key={empresa.codigo}>
                  <span>
                    {empresa.codigo} — {empresa.nombre}
                  </span>
                  <button
                    type="button"
                    className="btn-link"
                    disabled={disabled}
                    onClick={() => handleRemoveEmpresa(empresa.codigo)}
                  >
                    Quitar
                  </button>
                </li>
              ))}
            </ul>
          ) : null}

          <form className="admin-form" onSubmit={handleAddConcepto}>
            <h3>Nuevo concepto (pago terceros)</h3>
            <div className="admin-grid">
              <label className="field">
                <span>Grupo catálogo</span>
                <select
                  value={conceptoGrupo}
                  onChange={(e) => setConceptoGrupo(e.target.value as GrupoCatalogo)}
                  disabled={disabled}
                >
                  {GRUPOS_CATALOGO.map((grupo) => (
                    <option key={grupo} value={grupo}>
                      {grupo}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>Acreedor</span>
                <input
                  value={acreedor}
                  onChange={(e) => setAcreedor(e.target.value)}
                  disabled={disabled}
                  placeholder="Ej. 104744"
                />
              </label>
              <label className="field">
                <span>Concepto</span>
                <input
                  value={concepto}
                  onChange={(e) => setConcepto(e.target.value)}
                  disabled={disabled}
                  placeholder="Ej. 6MB6"
                />
              </label>
              <label className="field">
                <span>Descripción</span>
                <input
                  value={descripcion}
                  onChange={(e) => setDescripcion(e.target.value)}
                  disabled={disabled}
                  placeholder="Ej. AFC.BANCOLOMBIA"
                />
              </label>
              <label className="field">
                <span>Método de pago</span>
                <input
                  value={metodoPago}
                  onChange={(e) => setMetodoPago(e.target.value)}
                  disabled={disabled}
                  placeholder="Ej. PLANO"
                />
              </label>
            </div>
            <button type="submit" className="btn btn-primary" disabled={disabled}>
              Guardar concepto
            </button>
          </form>

          {customConceptos.length > 0 ? (
            <ul className="admin-list">
              {customConceptos.map((row) => (
                <li key={row.concepto}>
                  <span>
                    {row.concepto} · {row.acreedor} · {row.metodoPago}
                  </span>
                  <button
                    type="button"
                    className="btn-link"
                    disabled={disabled}
                    onClick={() => handleRemoveConcepto(row.concepto)}
                  >
                    Quitar
                  </button>
                </li>
              ))}
            </ul>
          ) : null}
        </div>
      ) : null}
    </div>
  )
}
