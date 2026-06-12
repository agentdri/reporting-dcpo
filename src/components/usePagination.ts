/**
 * ============================================================================
 * HOOK + CONFIG DE PAGINATION
 * ============================================================================
 *
 * Fichier dédié au hook `usePagination` et aux constantes globales.
 * Séparé du composant Pagination.tsx pour respecter la règle React-refresh
 * (un fichier qui exporte des composants ne doit exporter que des composants).
 *
 * Pattern d'utilisation :
 *
 *   const pagination = usePagination({
 *     total: filteredItems.length,
 *     resetKey: JSON.stringify(appliedFilters),
 *   })
 *
 *   const pagedItems = useMemo(
 *     () => filteredItems.slice(pagination.start, pagination.end),
 *     [filteredItems, pagination.start, pagination.end],
 *   )
 *
 * Voir Pagination.tsx pour le composant UI qui consomme ce hook.
 * ============================================================================
 */

import { useCallback, useState } from 'react'


/* ══════════════════════════════════════════════════════════════════════════
 * 🎛️  CONFIGURATION GLOBALE
 * ══════════════════════════════════════════════════════════════════════════
 *
 *   👉 POUR CHANGER LA PAGINATION DANS TOUTE L'APP, MODIFIER ICI 👈
 *
 * ══════════════════════════════════════════════════════════════════════════ */

/**
 * Taille de page par défaut quand une page ne précise rien.
 * Choisi à 50 : densité d'info élevée pour limiter les changements de page,
 * adaptée à un usage métier où l'utilisateur scanne souvent de grosses listes.
 *
 * Pour avoir un défaut différent : changer cette valeur OU passer
 * `initialPageSize: 25` aux options du hook depuis une page spécifique.
 */
export const DEFAULT_PAGE_SIZE = 50

/**
 * Tailles proposées dans le sélecteur "Afficher X par page".
 * Doit toujours contenir DEFAULT_PAGE_SIZE.
 *
 * Conseils :
 *   - Trop d'options → UX confuse
 *   - Trop peu → utilisateur frustré (impossible d'afficher 50/100)
 *   - 4-5 options est un bon équilibre
 */
export const PAGE_SIZE_OPTIONS: readonly number[] = [5, 10, 25, 50, 100]


/* ══════════════════════════════════════════════════════════════════════════
 * HOOK — usePagination
 * ══════════════════════════════════════════════════════════════════════════ */

/**
 * Options d'initialisation du hook.
 *
 *   - total          : nombre total d'éléments à paginer (longueur du tableau filtré)
 *   - initialPageSize: surcharge de DEFAULT_PAGE_SIZE pour cette page
 *   - resetKey       : valeur arbitraire qui, lorsqu'elle change, ramène
 *                      la pagination à la page 1. Typiquement passer une
 *                      string représentant l'état des filtres (ex:
 *                      JSON.stringify(appliedFilters)) → quand l'utilisateur
 *                      change un filtre, il revient automatiquement page 1.
 */
export interface UsePaginationOptions {
  total: number
  initialPageSize?: number
  resetKey?: unknown
}

/**
 * État retourné par le hook.
 *
 *   - page         : numéro de page courante (1-indexée, plus humain que 0)
 *   - pageSize     : nombre d'éléments par page
 *   - totalPages   : nombre total de pages (toujours >= 1, même si total=0)
 *   - start / end  : bornes de slice (start inclusif, end exclusif)
 *                    → `array.slice(start, end)` donne la page courante
 *   - setPage / setPageSize : setters bruts
 *   - goFirst/goPrev/goNext/goLast : navigation
 *   - canPrev/canNext : flags pour désactiver les boutons aux extrémités
 */
export interface PaginationState {
  page: number
  pageSize: number
  totalPages: number
  start: number
  end: number
  setPage: (page: number) => void
  setPageSize: (size: number) => void
  goFirst: () => void
  goPrev: () => void
  goNext: () => void
  goLast: () => void
  canPrev: boolean
  canNext: boolean
}

/**
 * Hook de gestion de pagination.
 *
 * Comportements clés :
 *   1. Page courante 1-indexée
 *   2. totalPages calculé via Math.ceil(total / pageSize), minimum 1
 *   3. Clamp automatique : si la page courante devient hors-limite
 *      (ex: filtrage réduit total), elle est ramenée à totalPages
 *   4. Reset automatique à la page 1 quand `resetKey` change
 *   5. Changer pageSize remet à la page 1 (sinon UX bizarre où l'utilisateur
 *      passe de 50/page à 10/page et se retrouve "perdu" loin du début)
 */
export function usePagination({
  total,
  initialPageSize = DEFAULT_PAGE_SIZE,
  resetKey,
}: UsePaginationOptions): PaginationState {
  const [page, setPageInternal] = useState(1)
  const [pageSize, setPageSizeInternal] = useState(initialPageSize)
  // Mémorise la dernière resetKey pour détecter les changements PENDANT le render
  // (pattern recommandé par React 19 pour éviter useEffect côté reset de state).
  const [prevResetKey, setPrevResetKey] = useState(resetKey)

  // ─── Reset à la page 1 sur changement de resetKey ──────────────────
  // Pattern "compute-during-render" (cf. https://react.dev/reference/react/useState
  // section "Storing information from previous renders") :
  // React détecte le setState pendant le render, jette le rendu en cours et
  // démarre un nouveau rendu avec la valeur à jour. Pas de cascade ni useEffect.
  if (!Object.is(prevResetKey, resetKey)) {
    setPrevResetKey(resetKey)
    setPageInternal(1)
  }

  // ─── Calculs dérivés ───────────────────────────────────────────────
  // Math.max(1, ...) : on garde au moins 1 page même si total=0
  const totalPages = Math.max(1, Math.ceil(total / pageSize))

  // ─── Clamp automatique si page hors-limite ─────────────────────────
  // Idem pattern compute-during-render : si le filtrage a réduit total
  // tel que la page courante n'existe plus, on la ramène à la dernière
  // page disponible. La condition garantit la convergence (page diminue,
  // donc au prochain render page <= totalPages → plus de setState).
  if (page > totalPages) {
    setPageInternal(totalPages)
  }

  const start = (page - 1) * pageSize
  const end = Math.min(start + pageSize, total)

  const canPrev = page > 1
  const canNext = page < totalPages

  // ─── Setters publics ───────────────────────────────────────────────
  const setPage = useCallback((p: number) => {
    setPageInternal(Math.max(1, Math.min(totalPages, Math.floor(p) || 1)))
  }, [totalPages])

  const setPageSize = useCallback((size: number) => {
    setPageSizeInternal(size)
    setPageInternal(1)
  }, [])

  // ─── Navigation ────────────────────────────────────────────────────
  const goFirst = useCallback(() => setPageInternal(1), [])
  const goPrev = useCallback(() => setPageInternal(p => Math.max(1, p - 1)), [])
  const goNext = useCallback(() => setPageInternal(p => Math.min(totalPages, p + 1)), [totalPages])
  const goLast = useCallback(() => setPageInternal(totalPages), [totalPages])

  return {
    page, pageSize, totalPages, start, end,
    setPage, setPageSize,
    goFirst, goPrev, goNext, goLast,
    canPrev, canNext,
  }
}
