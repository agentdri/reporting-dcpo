/**
 * ============================================================================
 * MODULE — DASHBOARDS POWER BI
 * ============================================================================
 *
 * Page générique qui embarque un rapport Power BI publié dans une iframe.
 *
 * Mode de publication :
 *   Les URLs `app.powerbi.com/view?r=...` sont des liens "Publier sur le
 *   web" — ils sont accessibles à TOUTE personne disposant de l'URL,
 *   sans authentification. À utiliser uniquement pour des dashboards
 *   non sensibles. Pour des dashboards à accès contrôlé, basculer sur
 *   "Embed for your organization" + Power BI Embedded SDK (plus complexe).
 *
 * Pourquoi un composant générique ?
 *   On a actuellement 3 dashboards (Numérisation, Créance hors bilan,
 *   Apurement Comex) avec exactement le même rendu (titre + iframe plein
 *   écran). Plutôt que dupliquer 3 fichiers, on paramétrise par les props.
 *   Pour en ajouter un 4e : créer juste une nouvelle entrée dans
 *   Dashboard.tsx → c'est tout.
 *
 * Layout :
 *   - Header titre (.content-header)
 *   - iframe qui occupe tout l'espace restant (flex: 1)
 *   - Bouton "Ouvrir dans un nouvel onglet" pour visualiser hors app si
 *     besoin (utile sur petit écran où l'iframe peut être trop tassée)
 *
 * Accessibilité :
 *   L'iframe a un attribut `title` obligatoire (sinon warning a11y).
 * ============================================================================
 */

export interface DashboardPowerBIProps {
  /** Titre affiché en haut de la page. */
  title: string
  /**
   * URL Power BI "view" publique (https://app.powerbi.com/view?r=...).
   * Doit être complète et valide ; sinon l'iframe affichera une page
   * d'erreur Power BI.
   */
  url: string
}

export default function DashboardPowerBI({ title, url }: DashboardPowerBIProps) {
  return (
    <>
      <div className="content-header" style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
        <h2 style={{ flex: 1, margin: 0 }}>{title}</h2>
        {/* Lien externe : ouvre le dashboard dans un nouvel onglet,
            utile pour profiter du plein écran natif Power BI. */}
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

      {/* L'iframe prend toute la hauteur restante de la zone de contenu.
          - calc(100vh - X) : on retire la topbar + le header pour que
            l'utilisateur n'ait pas à scroller verticalement.
          - allowFullScreen : permet l'icône plein écran Power BI native.
          - frameBorder=0 + style border: 'none' : double sécurité (l'attribut
            est déprécié en HTML5 mais certains navigateurs s'y réfèrent). */}
      <div
        style={{
          flex: 1,
          minHeight: 'calc(100vh - 180px)',
          marginTop: 16,
          background: '#f5f5f5',
          borderRadius: 8,
          overflow: 'hidden',
          border: '1px solid #e0e0e0',
        }}
      >
        <iframe
          title={title}
          src={url}
          style={{ width: '100%', height: '100%', minHeight: 'calc(100vh - 200px)', border: 'none', display: 'block' }}
          allowFullScreen
        />
      </div>
    </>
  )
}
