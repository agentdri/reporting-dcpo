/**
 * ============================================================================
 * COMPOSANT — <Pagination />
 * ============================================================================
 *
 * Barre de pagination réutilisable consommant un `PaginationState` produit
 * par le hook `usePagination` (cf. usePagination.ts).
 *
 * Pourquoi composant séparé du hook ?
 *   - Permet à une page d'afficher la pagination en haut ET en bas
 *     (les deux composants partagent le même state via la même instance
 *     du hook)
 *   - Respect règle React-refresh : un fichier qui exporte des composants
 *     ne doit exporter que des composants (le hook + constantes sont dans
 *     usePagination.ts)
 * ============================================================================
 */

import { type PaginationState, PAGE_SIZE_OPTIONS } from './usePagination'


/**
 * Props du composant Pagination.
 *
 *   - state                : retour du hook usePagination (source de vérité)
 *   - total                : nombre total d'éléments (filtrés) pour affichage
 *                            "X-Y sur N {label}"
 *   - itemLabel            : nom pluriel des éléments (ex: "anomalies",
 *                            "rapports", "agents") pour l'affichage du compteur
 *   - showPageSizeSelector : true par défaut, à false pour cacher le select
 *                            (utile si on veut une taille fixe sans contrôle UX)
 *   - pageSizeOptions      : surcharge la liste des options du select
 */
export interface PaginationProps {
  state: PaginationState
  total: number
  itemLabel?: string
  showPageSizeSelector?: boolean
  pageSizeOptions?: readonly number[]
}

/**
 * Barre de pagination affichant :
 *   [ 1-10 sur 87 anomalies ] [« ‹ Page 1 / 9 › »] [Afficher [10▾] par page]
 *
 * Comportements :
 *   - Si total === 0 : on ne rend rien (table déjà vide, pas besoin de UI
 *     supplémentaire qui afficherait "0-0 sur 0")
 *   - Les boutons sont désactivés (disabled + opacity 0.4) aux extrémités
 *     pour signaler visuellement les bornes
 *   - Le select de taille rappelle systématiquement la valeur actuelle de
 *     state.pageSize (binding contrôlé)
 *
 * Accessibilité :
 *   - aria-label sur chaque bouton (« Première page », « Page précédente »...)
 *   - Le `disabled` natif rend les boutons non-focusable au clavier quand
 *     ils sont aux bornes
 */
export function Pagination({
  state,
  total,
  itemLabel = 'éléments',
  showPageSizeSelector = true,
  pageSizeOptions = PAGE_SIZE_OPTIONS,
}: PaginationProps) {
  // Court-circuit : pas d'éléments → pas de barre
  if (total === 0) return null

  const {
    page, totalPages, pageSize, start, end,
    goFirst, goPrev, goNext, goLast,
    canPrev, canNext, setPageSize,
  } = state

  return (
    <div className="pagination" role="navigation" aria-label="Pagination">
      {/* ─── Zone gauche : compteur "X-Y sur N items" ──────────────── */}
      {/* start est 0-indexé en interne ; on affiche +1 pour l'humain */}
      <div className="pagination-info">
        <span>
          <strong>{start + 1}-{end}</strong> sur <strong>{total}</strong> {itemLabel}
        </span>
      </div>

      {/* ─── Zone centre : boutons de navigation ─────────────────── */}
      <div className="pagination-controls">
        <button
          type="button"
          disabled={!canPrev}
          onClick={goFirst}
          aria-label="Première page"
          title="Première page"
        >
          «
        </button>
        <button
          type="button"
          disabled={!canPrev}
          onClick={goPrev}
          aria-label="Page précédente"
          title="Page précédente"
        >
          ‹
        </button>
        <span className="pagination-page">
          Page <strong>{page}</strong> / {totalPages}
        </span>
        <button
          type="button"
          disabled={!canNext}
          onClick={goNext}
          aria-label="Page suivante"
          title="Page suivante"
        >
          ›
        </button>
        <button
          type="button"
          disabled={!canNext}
          onClick={goLast}
          aria-label="Dernière page"
          title="Dernière page"
        >
          »
        </button>
      </div>

      {/* ─── Zone droite : sélecteur de taille de page ──────────── */}
      {showPageSizeSelector && (
        <div className="pagination-size">
          <label htmlFor="pagination-page-size">Afficher</label>
          <select
            id="pagination-page-size"
            value={pageSize}
            onChange={e => setPageSize(Number(e.target.value))}
          >
            {pageSizeOptions.map(s => (
              <option key={s} value={s}>{s}</option>
            ))}
          </select>
          <span>par page</span>
        </div>
      )}
    </div>
  )
}
