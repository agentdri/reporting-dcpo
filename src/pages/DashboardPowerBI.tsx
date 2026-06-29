/**
 * ============================================================================
 * MODULE — DASHBOARDS POWER BI
 * ============================================================================
 *
 * Page générique qui embarque un rapport Power BI publié.
 *
 * APPROCHE SIMPLE & ROBUSTE
 * -------------------------
 * On rend simplement l'iframe avec les attributs nécessaires à Power BI.
 *   - PAS d'auto-détection de blocage (la détection par timer coupait
 *     l'affichage en local quand Power BI prenait juste du temps à charger
 *     son bundle JS d'environ 5 Mo).
 *   - L'iframe est TOUJOURS rendu. Si la CSP du host le bloque, le browser
 *     affiche son propre placeholder à l'intérieur — c'est OK, on a quand
 *     même un bouton "Ouvrir dans un nouvel onglet" en permanence.
 *   - Spinner overlay léger pendant le chargement initial, masqué dès que
 *     `onLoad` se déclenche.
 *
 * ATTRIBUTS IFRAME
 * ----------------
 *   - allow="fullscreen *; clipboard-read; clipboard-write" :
 *       active les features modernes Power BI (plein écran, copier-coller)
 *   - referrerPolicy="no-referrer-when-downgrade" :
 *       compatible avec les exigences de Power BI sur le referrer
 *   - PAS de `sandbox` : Power BI a besoin d'exécuter son JS + cookies tiers
 *     pour l'auth. Un sandbox même permissif casse l'iframe.
 *
 * EN POWER APPS PUBLIÉ
 * --------------------
 * Si la CSP du tenant Power Apps n'inclut pas `app.powerbi.com` dans
 * `frame-src`, le browser affichera son propre message de blocage dans
 * l'iframe. Solutions :
 *   - Demander à l'admin Power Platform de whitelist app.powerbi.com
 *     (voir doc Microsoft "Content security policy for canvas apps")
 *   - Côté UX : le bouton "Ouvrir dans un nouvel onglet" reste utilisable
 *     en permanence (un <a target="_blank"> n'est PAS soumis à la CSP).
 *
 * Pour activer l'authentification utilisateur sur le dashboard (au lieu de
 * "publish to web" public), basculer sur Power BI Embedded + SDK
 * powerbi-client. Plus complexe mais 100% intégré.
 * ============================================================================
 */

import { useState } from 'react'

export interface DashboardPowerBIProps {
  /** Titre affiché en haut de la page (et attribut title de l'iframe). */
  title: string
  /**
   * URL Power BI "view" publique (https://app.powerbi.com/view?r=...).
   * Si invalide, Power BI affiche sa propre page d'erreur dans l'iframe.
   */
  url: string
}

export default function DashboardPowerBI({ title, url }: DashboardPowerBIProps) {
  // Spinner masqué dès que l'iframe émet onLoad. Pas de détection de
  // blocage : on accepte que le iframe puisse mettre du temps ou être
  // bloqué par CSP — dans tous les cas, le bouton "Ouvrir nouvel onglet"
  // reste disponible.
  const [loaded, setLoaded] = useState(false)

  return (
    <>
      {/* ─── Header titre + CTA externe (toujours visible) ──────────────── */}
      <div
        className="content-header"
        style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}
      >
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

      {/* ─── Zone iframe ────────────────────────────────────────────────── */}
      <div className="powerbi-embed-wrap">
        {/* key={url} : remonte un iframe FRESH à chaque changement de
            dashboard (sinon onLoad d'un ancien iframe contamine le nouveau). */}
        <iframe
          key={url}
          title={title}
          src={url}
          onLoad={() => setLoaded(true)}
          allow="fullscreen *; clipboard-read; clipboard-write"
          referrerPolicy="no-referrer-when-downgrade"
          allowFullScreen
          style={{
            width: '100%',
            height: 'calc(100vh - 200px)',
            border: 'none',
            display: 'block',
          }}
        />

        {!loaded && (
          <div className="powerbi-embed-loading" role="status" aria-live="polite">
            <div className="app-loading-spinner" aria-hidden="true" />
            <span>Chargement du tableau de bord…</span>
          </div>
        )}
      </div>
    </>
  )
}
