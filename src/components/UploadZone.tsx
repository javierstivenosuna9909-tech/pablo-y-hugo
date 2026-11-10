import { useRef, useState, type DragEvent, type ChangeEvent } from 'react'

type UploadZoneProps = {
  file: File | null
  disabled?: boolean
  onFileSelected: (file: File | null) => void
  title?: string
  hint?: string
}

export function UploadZone({
  file,
  disabled,
  onFileSelected,
  title = 'Arrastra el Excel aquí',
  hint = 'o haz clic para seleccionar un .xlsx con NOMINA (y CXP si ya viene)',
}: UploadZoneProps) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [dragging, setDragging] = useState(false)

  const acceptFile = (next: File | undefined | null) => {
    if (!next) return
    if (!next.name.toLowerCase().endsWith('.xlsx')) {
      onFileSelected(null)
      return
    }
    onFileSelected(next)
  }

  const onDrop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault()
    setDragging(false)
    if (disabled) return
    acceptFile(event.dataTransfer.files?.[0])
  }

  const onChange = (event: ChangeEvent<HTMLInputElement>) => {
    acceptFile(event.target.files?.[0])
    event.target.value = ''
  }

  return (
    <div
      className={`upload-zone${dragging ? ' is-dragging' : ''}${disabled ? ' is-disabled' : ''}`}
      onDragEnter={(e) => {
        e.preventDefault()
        if (!disabled) setDragging(true)
      }}
      onDragOver={(e) => e.preventDefault()}
      onDragLeave={(e) => {
        e.preventDefault()
        setDragging(false)
      }}
      onDrop={onDrop}
      onClick={() => !disabled && inputRef.current?.click()}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if ((e.key === 'Enter' || e.key === ' ') && !disabled) {
          e.preventDefault()
          inputRef.current?.click()
        }
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        hidden
        onChange={onChange}
        disabled={disabled}
      />
      <p className="upload-title">{title}</p>
      <p className="upload-hint">{hint}</p>
      {file ? <p className="upload-file">{file.name}</p> : null}
    </div>
  )
}
