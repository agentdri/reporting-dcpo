/**
 * Overlay de modale rendu via portail sur document.body.
 *
 * Les modales étaient auparavant dans .dashboard-content, dont GSAP applique
 * un transform (useFadeUp) : cela recréait un contexte d'empilement et
 * position:fixed ne couvrait plus tout l'écran (modale coupée par la sidebar).
 */
import { useEffect, type ReactNode } from 'react'
import { createPortal } from 'react-dom'

interface ModalOverlayProps {
  onClose: () => void
  children: ReactNode
  /** Classes additionnelles (ex. bulletin-modal-overlay). */
  className?: string
}

export function ModalOverlay({ onClose, children, className }: ModalOverlayProps) {
  useEffect(() => {
    const prev = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    return () => {
      document.body.style.overflow = prev
    }
  }, [])

  const classes = className ? `modal-overlay ${className}` : 'modal-overlay'

  return createPortal(
    <div className={classes} onClick={onClose} role="presentation">
      {children}
    </div>,
    document.body,
  )
}
