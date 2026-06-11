/**
 * ============================================================================
 * MODULE — DASHBOARDS POWER BI
 * ============================================================================
 *
 * Page générique qui embarque un rapport Power BI publié.
 *
 * PROBLÉMATIQUE : EMBEDDING DANS POWER APPS
 * -----------------------------------------
 * L'app est elle-même hébergée par Power Apps (`apps.powerapps.com`) dans
 * un iframe. Quand on tente d'embarquer un autre iframe (Power BI) à
 * l'intérieur, la CSP (Content Security Policy) imposée par le shell Power
 * Apps peut BLOQUER le chargement (directive `frame-src` qui n'inclut pas
 * `app.powerbi.com` par défaut).
 *
 * En local (`npm run dev` sur localhost) : pas de CSP → l'iframe marche.
 * Sur l'app publiée : la CSP s'applique → l'iframe peut afficher un
 * placeholder "contenu bloqué".
 *
 * STRATÉGIE DE FALLBACK AUTOMATIQUE
 * ---------------------------------
 * On rend le iframe ET on observe son chargement :
 *   1. État initial : 'loading' (un spinner + bouton CTA disponible)
 *   2. Si `onLoad` se déclenche → état 'loaded' (l'iframe est visible)
 *   3. Si après LOAD_TIMEOUT_MS rien ne se passe → on déduit que c'est
 *      bloqué (CSP / X-Frame-Options) → on bascule en 'blocked' et on
 *      masque l'iframe au profit d'un panneau CTA stylisé qui invite
 *      explicitement à ouvrir dans un nouvel onglet.
 *
 * NB : `onError` sur un iframe est peu fiable cross-origin (le browser ne
 * remonte pas l'erreur de CSP au parent). Le timer est le mécanisme le
 * plus universel pour détecter un blocage silencieux.
 *
 * ATTRIBUTS IFRAME PERMISSIFS
 * ---------------------------
 *   - allow="fullscreen *; clipboard-read; clipboard-write" :
 *       active les features modernes Power BI (plein écran, copier-coller
 *       de visuels) — utile sur les configurations CSP qui distinguent
 *       chargement et permissions
 *   - referrerPolicy="no-referrer-when-downgrade" :
 *       compatible Power BI (sinon il peut refuser le referrer "null")
 *   - PAS de `sandbox` : ajouter `sandbox=...` bloque Power BI qui a besoin
 *     d'exécuter ses propres scripts + cookies tiers (auth Microsoft).
 *
 * BOUTON "OUVRIR DANS UN NOUVEL ONGLET"
 * ------------------------------------
 * Visible en permanence (état loading, loaded ET blocked) — c'est le filet
 * de sécurité ultime garanti par le navigateur (un `<a target="_blank">`
 * n'est PAS soumis à la CSP frame-src du parent).
 * ============================================================================
 */

import { useEffect, useRef, useState } from 'react'

/** Délai au-delà duquel on considère que l'iframe est bloqué (ms). */
const LOAD_TIMEOUT_MS = 4000

export interface DashboardPowerBIProps {
  /** Titre affiché en haut de la page (et utilisé pour l'attribut iframe title). */
  title: string
  /**
   * URL Power BI "view" publique (https://app.powerbi.com/view?r=...).
   * Doit être complète et valide ; sinon l'iframe affichera une page
   * d'erreur Power BI (et passera quand même en 'loaded').
   */
  url: string
}

type IframeStatus = 'loading' | 'loaded' | 'blocked'

export default function DashboardPowerBI({ title, url }: DashboardPowerBIProps) {
  const [status, setStatus] = useState<IframeStatus>('loading')
  /**
   * Pattern "compute during render" (React 19) pour reset le state quand
   * l'URL change (navigation entre dashboards). On compare l'URL précédente
   * stockée en state à l'URL courante — si différentes, on remet à 'loading'
   * pendant le rendu, ce qui évite un `setState` synchrone dans useEffect
   * (interdit par le linter react-hooks/set-state-in-effect en React 19).
   */
  const [prevUrl, setPrevUrl] = useState(url)
  if (prevUrl !== url) {
    setPrevUrl(url)
    setStatus('loading')
  }

  // Ref vers le timer : permet de l'annuler proprement si l'iframe charge
  // avant LOAD_TIMEOUT_MS, ou si le composant se démonte (changement
  // d'onglet, par exemple).
  const timerRef = useRef<number | undefined>(undefined)

  useEffect(() => {
    // L'effet relance le timer à chaque changement d'URL. Le state 'status'
    // a déjà été remis à 'loading' par le compute-during-render ci-dessus,
    // donc on a juste à (re)programmer le timeout de détection blocage.
    timerRef.current = window.setTimeout(() => {
      // Si on est toujours en 'loading' au déclenchement → l'iframe n'a
      // pas appelé onLoad. C'est presque toujours un blocage CSP côté
      // host (Power Apps refuse de frame Power BI).
      // Si onLoad est arrivé entre-temps, handleLoad aura déjà annulé
      // ce timer (donc on n'arrive jamais ici).
      setStatus(prev => (prev === 'loading' ? 'blocked' : prev))
    }, LOAD_TIMEOUT_MS)
    return () => {
      if (timerRef.current !== undefined) window.clearTimeout(timerRef.current)
    }
  }, [url])

  /** Appelé par l'iframe quand le contenu termine de charger. */
  const handleLoad = () => {
    // On annule le timer "blocked" pour éviter une fausse bascule
    if (timerRef.current !== undefined) {
      window.clearTimeout(timerRef.current)
      timerRef.current = undefined
    }
    setStatus('loaded')
  }

  return (
    <>
      {/* ─── Header titre + CTA externe ────────────────────────────────────
          Le bouton "Ouvrir dans un nouvel onglet" reste visible en
          permanence — c'est le filet de sécurité même quand l'iframe
          charge correctement (utilisateur qui préfère le plein écran). */}
      <div className="content-header" style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <h2 style={{ flex: 1, margin: 0 }}>{title}</h2>
        <a
          href={url}
          target="_blank"
          rel="noreferrer"
          className="btn-add"
          style={{ textDecoration: 'none' }}
        >
          ↗ Ouvrir dans un nouvel onglet
        </a>
      </div>

      {/* ─── Zone d'affichage : iframe OU panneau "bloqué" ────────────────
          Hauteur calculée pour remplir la zone restante sous la topbar +
          le content-header (~180 px). */}
      <div
        style={{
          flex: 1,
          minHeight: 'calc(100vh - 200px)',
          marginTop: 16,
          background: '#f5f5f5',
          borderRadius: 8,
          overflow: 'hidden',
          border: '1px solid #e0e0e0',
          position: 'relative',
        }}
      >
        {status !== 'blocked' && (
          <iframe
            // key forcé sur l'URL : remonte un iframe FRESH à chaque
            // changement de dashboard (évite que onLoad d'un ancien iframe
            // contamine la détection du nouveau).
            key={url}
            title={title}
            src={url}
            onLoad={handleLoad}
            allow="fullscreen *; clipboard-read; clipboard-write"
            referrerPolicy="no-referrer-when-downgrade"
            style={{
              width: '100%',
              height: '100%',
              minHeight: 'calc(100vh - 220px)',
              border: 'none',
              display: 'block',
              // Tant qu'on est en loading on cache visuellement l'iframe
              // pour éviter le flash gris vide pendant la détection.
              visibility: status === 'loaded' ? 'visible' : 'hidden',
            }}
            allowFullScreen
          />
        )}

        {/* ─── Overlay "loading" (3-4 sec max) ────────────────────────── */}
        {status === 'loading' && (
          <div
            style={{
              position: 'absolute', inset: 0,
              display: 'flex', flexDirection: 'column',
              alignItems: 'center', justifyContent: 'center',
              gap: 12, color: '#666', fontSize: 14,
            }}
          >
            <div
              style={{
                width: 36, height: 36, borderRadius: '50%',
                border: '3px solid #e0e0e0', borderTopColor: '#e63946',
                animation: 'spin 0.8s linear infinite',
              }}
              aria-hidden="true"
            />
            <span>Chargement du tableau de bord…</span>
          </div>
        )}

        {/* ─── Fallback "bloqué" : panneau CTA ──────────────────────────
            Affiché quand le timer s'écoule sans onLoad → presque toujours
            une CSP frame-src côté Power Apps qui n'autorise pas Power BI.
            Plutôt que d'afficher un iframe gris vide, on explique
            clairement le pourquoi + on propose l'action principale.        */}
        {status === 'blocked' && (
          <div
            style={{
              position: 'absolute', inset: 0,
              display: 'flex', flexDirection: 'column',
              alignItems: 'center', justifyContent: 'center',
              padding: 32, textAlign: 'center', gap: 12,
            }}
          >
            <div style={{ fontSize: 48 }} aria-hidden="true">🔒</div>
            <h3 style={{ margin: 0, color: '#1a1a2e', fontSize: 18 }}>
              Affichage embarqué non disponible
            </h3>
            <p style={{ margin: 0, color: '#555', maxWidth: 540, lineHeight: 1.5 }}>
              L'environnement Power Apps n'autorise pas l'embarquement direct
              de ce tableau de bord Power BI. Ouvrez-le dans un nouvel onglet
              pour consulter le contenu en plein écran.
            </p>
            <a
              href={url}
              target="_blank"
              rel="noreferrer"
              className="btn-add"
              style={{
                textDecoration: 'none',
                marginTop: 8,
                padding: '10px 20px',
                fontSize: 14,
              }}
            >
              ↗ Ouvrir le tableau de bord
            </a>
            <p style={{ margin: '12px 0 0', fontSize: 12, color: '#888' }}>
              Astuce : épinglez l'onglet pour y revenir rapidement.
            </p>
          </div>
        )}
      </div>

      {/* Keyframes pour l'animation du spinner — injection inline pour ne
          pas dépendre d'un fichier CSS séparé sur cette petite page. */}
      <style>{`
        @keyframes spin {
          to { transform: rotate(360deg); }
        }
      `}</style>
    </>
  )
}
