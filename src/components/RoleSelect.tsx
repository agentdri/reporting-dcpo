/**
 * Sélecteur de rôle (topbar) — remplace le <select> natif dont le menu
 * déroulant hérite du style système (souvent bleu) et casse la charte.
 */
import { useEffect, useRef, useState } from 'react'
import NavIcon from './NavIcon'

export interface RoleSelectOption {
  value: string
  label: string
}

interface RoleSelectProps {
  value: string
  options: RoleSelectOption[]
  onChange: (value: string) => void
  'aria-label'?: string
}

export function RoleSelect({ value, options, onChange, 'aria-label': ariaLabel }: RoleSelectProps) {
  const [open, setOpen] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    if (!open) return
    const onOutside = (e: MouseEvent) => {
      if (!rootRef.current?.contains(e.target as Node)) setOpen(false)
    }
    const onEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false)
    }
    document.addEventListener('mousedown', onOutside)
    document.addEventListener('keydown', onEscape)
    return () => {
      document.removeEventListener('mousedown', onOutside)
      document.removeEventListener('keydown', onEscape)
    }
  }, [open])

  const current = options.find(o => o.value === value)

  return (
    <div className="role-select" ref={rootRef}>
      <button
        type="button"
        className="role-select-trigger"
        onClick={() => setOpen(v => !v)}
        aria-haspopup="listbox"
        aria-expanded={open}
        aria-label={ariaLabel}
      >
        <span className="role-select-icon" aria-hidden="true">
          <NavIcon name="user" size={14} />
        </span>
        <span className="role-select-value">{current?.label ?? value}</span>
        <span className="role-select-chevron" aria-hidden="true" />
      </button>
      {open && (
        <ul className="role-select-menu" role="listbox" aria-label={ariaLabel}>
          {options.map(opt => (
            <li key={opt.value} role="presentation">
              <button
                type="button"
                role="option"
                aria-selected={opt.value === value}
                className={`role-select-option${opt.value === value ? ' is-selected' : ''}`}
                onClick={() => {
                  onChange(opt.value)
                  setOpen(false)
                }}
              >
                {opt.label}
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  )
}
