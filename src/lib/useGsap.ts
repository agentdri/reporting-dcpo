/**
 * ============================================================================
 * useGsap — Hooks d'animation d'entrée réutilisables (GSAP)
 * ============================================================================
 *
 * Centralise les animations « charte modernisée » du template :
 *   - useFadeUp  : fondu + translation verticale d'un élément (contenu de page)
 *   - useStagger : apparition en cascade des enfants (cartes stats, lignes de
 *                  tableau) au chargement des données
 *   - useSlideIn : glissement latéral d'un panneau/drawer à l'ouverture
 *
 * Chaque hook :
 *   - utilise useLayoutEffect (évite tout flash avant la pose de l'état initial) ;
 *   - encapsule le tween dans gsap.matchMedia('(prefers-reduced-motion:
 *     no-preference)') : les utilisateurs sensibles au mouvement n'ont AUCUNE
 *     animation (l'élément reste à son état final naturel) ;
 *   - nettoie via mm.revert() au démontage / changement de dépendances.
 *
 * Les @keyframes CSS (fadeUp/slideIn/overlayIn de theme.css) restent comme
 * fallback no-JS ; ces hooks les remplacent dynamiquement côté React.
 * ============================================================================
 */

import { useLayoutEffect, type RefObject, type DependencyList } from 'react'
import { gsap } from 'gsap'

/** Media query : on n'anime QUE si l'utilisateur n'a pas demandé moins de mouvement. */
const MOTION_OK = '(prefers-reduced-motion: no-preference)'

/**
 * Fondu + montée d'un élément au montage (et à chaque changement de `deps`).
 * Idéal pour la zone de contenu : passer [activeTab] rejoue l'entrée à chaque
 * changement d'onglet.
 *
 * @param ref  élément cible
 * @param deps dépendances qui relancent l'animation (défaut : montage seul)
 */
export function useFadeUp(
  ref: RefObject<HTMLElement | null>,
  deps: DependencyList = [],
) {
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const mm = gsap.matchMedia()
    mm.add(MOTION_OK, () => {
      gsap.from(el, { opacity: 0, y: 10, duration: 0.35, ease: 'power2.out' })
    })
    return () => mm.revert()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps)
}

interface StaggerOptions {
  /** Sélecteur CSS des enfants à animer (ex: 'tbody tr', '.stat-card'). */
  selector: string
  /** Dépendances qui relancent la cascade (typiquement les données chargées). */
  deps?: DependencyList
  /** Délai entre chaque enfant (s). Défaut 0.05. */
  stagger?: number
}

/**
 * Apparition en cascade des enfants `selector` du conteneur `ref`.
 * Utile après chargement asynchrone d'une liste : passer la donnée en `deps`.
 */
export function useStagger(
  ref: RefObject<HTMLElement | null>,
  { selector, deps = [], stagger = 0.05 }: StaggerOptions,
) {
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return
    const targets = el.querySelectorAll(selector)
    if (!targets.length) return
    const mm = gsap.matchMedia()
    mm.add(MOTION_OK, () => {
      gsap.from(targets, {
        opacity: 0,
        y: 12,
        duration: 0.3,
        stagger,
        ease: 'power2.out',
      })
    })
    return () => mm.revert()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, deps)
}

/**
 * Animation d'entrée de la sidebar (au montage du Dashboard).
 *
 * Orchestrée en timeline GSAP : le logo glisse depuis la gauche, puis les
 * entrées de navigation apparaissent en cascade (slide-in latéral), et enfin le
 * bloc utilisateur en pied monte en fondu. Donne le « réveil » progressif de la
 * navigation demandé dans la charte modernisée.
 *
 * Les cibles sont résolues par sélecteur DANS le `<nav>` passé en `ref`, donc le
 * hook reste indépendant du markup interne. Respecte prefers-reduced-motion via
 * matchMedia (aucune animation pour les utilisateurs sensibles au mouvement).
 *
 * @param ref élément racine de la sidebar (le `<nav class="dashboard-nav">`)
 */
export function useNavReveal(ref: RefObject<HTMLElement | null>) {
  useLayoutEffect(() => {
    const el = ref.current
    if (!el) return

    // Cibles résolues une seule fois. Le logo, les entrées de premier niveau
    // (enfants directs de .nav-list) et le bloc utilisateur. On NE touche pas
    // aux sous-items des groupes pour garder une cascade propre.
    const logo = el.querySelector('.nav-logo')
    const items = Array.from(el.querySelectorAll<HTMLElement>('.nav-list > *'))
    const user = el.querySelector('.nav-user')
    const targets = [logo, ...items, user].filter(Boolean) as Element[]

    // Filet de sécurité : ramène TOUJOURS les cibles à leur état CSS visible
    // (supprime opacity/transform inline posés par GSAP). Indispensable car un
    // `from` interrompu — typiquement le double-montage de React.StrictMode en
    // dev — laisserait sinon ces éléments à opacity:0 (invisibles).
    const forceVisible = () => {
      if (targets.length) gsap.set(targets, { clearProps: 'all' })
    }

    const mm = gsap.matchMedia()
    mm.add(MOTION_OK, () => {
      const tl = gsap.timeline({
        defaults: { ease: 'power3.out' },
        // clearProps à la fin de chaque tween → aucun style inline résiduel.
        onComplete: forceVisible,
      })
      if (logo) tl.from(logo, { autoAlpha: 0, x: -16, duration: 0.4 }, 0)
      if (items.length)
        tl.from(items, { autoAlpha: 0, x: -14, duration: 0.35, stagger: 0.05 }, 0.12)
      if (user) tl.from(user, { autoAlpha: 0, y: 10, duration: 0.4 }, '-=0.2')
      // Cleanup de la condition matchMedia : tue le tween et force la visibilité
      // si on est démonté/reverti en plein vol.
      return () => {
        tl.kill()
        forceVisible()
      }
    })

    return () => {
      mm.revert()
      forceVisible()
    }
  }, [ref])
}

/**
 * Glissement latéral (droite → place) d'un panneau/drawer quand `isOpen` devient
 * vrai. À brancher sur le ref du panneau monté conditionnellement.
 */
export function useSlideIn(
  ref: RefObject<HTMLElement | null>,
  isOpen: boolean,
) {
  useLayoutEffect(() => {
    const el = ref.current
    if (!el || !isOpen) return
    const mm = gsap.matchMedia()
    mm.add(MOTION_OK, () => {
      gsap.fromTo(
        el,
        { xPercent: 100 },
        { xPercent: 0, duration: 0.3, ease: 'power3.out' },
      )
    })
    return () => mm.revert()
  }, [ref, isOpen])
}
