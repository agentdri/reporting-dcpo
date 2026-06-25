/**
 * ============================================================================
 * NavIcon — Jeu d'icônes SVG de la sidebar (style « ligne », charte modernisée)
 * ============================================================================
 *
 * Remplace les emojis de la navigation par des icônes vectorielles cohérentes
 * (trait de 2px, viewBox 24×24, `stroke="currentColor"`) reprises de la maquette
 * « Interface Modernisée ». Comme le trait suit `currentColor`, chaque icône
 * hérite automatiquement de la couleur du lien (état hover / actif gérés en CSS),
 * sans aucune règle de couleur dédiée.
 *
 * Chaque entrée de ICONS est une liste de « primitives » :
 *   - une chaîne  → rendue comme <path d="…" />
 *   - un tuple    → [nomElement, attributs] (ex: ['circle', { cx, cy, r }])
 *
 * Usage :
 *   <NavIcon name="alert" />          // taille par défaut 16px
 *   <NavIcon name="users" size={18} />
 */
import { createElement } from 'react'

/** Primitive d'icône : soit un `d` de <path>, soit [element, attributs]. */
type IconPart = string | [string, Record<string, number>]

/**
 * Table des icônes. Les clés sont référencées par `icon` dans la nav du
 * Dashboard. Tracés repris du jeu d'icônes « line » de la maquette.
 */
const ICONS: Record<string, IconPart[]> = {
  home: ['M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z', 'M9 22V12h6v10'],
  alert: ['M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0zM12 9v4M12 17h.01'],
  list: ['M8 6h13', 'M8 12h13', 'M8 18h13', 'M3 6h.01', 'M3 12h.01', 'M3 18h.01'],
  file: ['M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z', 'M14 2v6h6', 'M16 13H8', 'M16 17H8', 'M10 9H8'],
  users: ['M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2', ['circle', { cx: 9, cy: 7, r: 4 }], 'M23 21v-2a4 4 0 0 0-3-3.87', 'M16 3.13a4 4 0 0 1 0 7.75'],
  user: ['M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2', ['circle', { cx: 12, cy: 7, r: 4 }]],
  check: ['M20 6L9 17l-5-5'],
  target: [['circle', { cx: 12, cy: 12, r: 10 }], ['circle', { cx: 12, cy: 12, r: 6 }], ['circle', { cx: 12, cy: 12, r: 2 }]],
  tool: ['M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z'],
  bar: ['M18 20V10', 'M12 20V4', 'M6 20v-6'],
  settings: [['circle', { cx: 12, cy: 12, r: 3 }], 'M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1-2.83 2.83l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-4 0v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83-2.83l.06-.06A1.65 1.65 0 0 0 4.68 15a1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1 0-4h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 2.83-2.83l.06.06A1.65 1.65 0 0 0 9 4.68a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 4 0v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 2.83l-.06.06A1.65 1.65 0 0 0 19.4 9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 0 4h-.09a1.65 1.65 0 0 0-1.51 1z'],
  calendar: [['rect', { x: 3, y: 4, width: 18, height: 18, rx: 2 }], 'M16 2v4', 'M8 2v4', 'M3 10h18'],
  edit: ['M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7', 'M18.5 2.5a2.12 2.12 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z'],
  folder: ['M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z'],
  building: ['M3 21h18', 'M5 21V7l8-4v18', 'M19 21V11l-6-4', 'M9 9h.01', 'M9 13h.01', 'M9 17h.01'],
  money: ['M12 1v22', 'M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6'],
  trend: ['M23 6l-9.5 9.5-5-5L1 18', 'M17 6h6v6'],
  link: ['M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71', 'M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71'],
}

interface NavIconProps {
  /** Clé de l'icône (cf. table ICONS). */
  name: string
  /** Taille en pixels (carré). Défaut 16, comme la maquette. */
  size?: number
}

/**
 * Rend l'icône SVG `name`. Renvoie `null` si la clé est inconnue (l'item
 * s'affiche alors sans icône plutôt que de casser le rendu).
 */
export default function NavIcon({ name, size = 16 }: NavIconProps) {
  const parts = ICONS[name]
  if (!parts) return null
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={2}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {parts.map((p, i) =>
        typeof p === 'string'
          ? createElement('path', { key: i, d: p })
          : createElement(p[0], { key: i, ...p[1] }),
      )}
    </svg>
  )
}
