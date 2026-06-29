/**
 * ============================================================================
 * COMPOSANT — <ProgressBar />
 * ============================================================================
 *
 * Barre de progression simple, utilisée dans les tableaux pour visualiser un
 * pourcentage d'avancement (ex: taux d'évolution d'un contrôle).
 *
 * - Couleur automatique selon le seuil :
 *     < 30 %  → rouge (alerte)
 *     < 70 %  → orange (en cours)
 *     ≥ 70 %  → vert (objectif atteint ou proche)
 *
 * - Affichage du pourcentage centré sur la barre, avec contraste fort
 *   pour rester lisible quel que soit le remplissage.
 * ============================================================================
 */

export interface ProgressBarProps {
  /** Valeur en pourcentage (0–100). Tronquée si hors bornes. */
  value: number
  /** Optionnel : texte alternatif affiché à la place du % (ex: "3/12"). */
  label?: string
  /** Optionnel : titre/info-bulle pour info supplémentaire au survol. */
  title?: string
  /**
   * Optionnel : couleur de remplissage à forcer (CSS color).
   * Si fourni, désactive la coloration automatique par seuil.
   * Utile quand le statut/sens visuel ne dépend pas du % (ex: indicateur
   * neutre ou toujours positif sur une page donnée).
   */
  color?: string
}

/** Détermine la couleur de remplissage selon le seuil (couleur automatique). */
function colorForValue(pct: number): string {
  if (pct < 30) return 'var(--rdcpo-red)'
  if (pct < 70) return 'var(--rdcpo-orange)'
  return 'var(--rdcpo-green)'
}

export function ProgressBar({ value, label, title, color }: ProgressBarProps) {
  const pct = Math.max(0, Math.min(100, Math.round(value)))
  // Priorité à la couleur fournie en prop ; sinon couleur automatique.
  const fill = color ?? colorForValue(pct)
  return (
    <div className="progress-bar" title={title ?? `${pct}%`} role="progressbar" aria-valuenow={pct} aria-valuemin={0} aria-valuemax={100}>
      <div
        className="progress-bar-fill"
        style={{ width: `${pct}%`, background: fill }}
        aria-hidden="true"
      />
      <span className="progress-bar-label">{label ?? `${pct}%`}</span>
    </div>
  )
}
