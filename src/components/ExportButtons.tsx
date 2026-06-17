/**
 * ============================================================================
 * EXPORT BUTTONS — bouton groupé "Excel + PDF"
 * ============================================================================
 *
 * Composant réutilisable que chaque page à tableau peut placer dans sa
 * barre d'outils. Encapsule l'appel aux helpers de `src/lib/exporters.ts`.
 *
 * Les données sont fournies par la page via les callbacks `getRows` /
 * `getHeaders` (plutôt que via des props directes), pour que l'évaluation
 * soit faite AU MOMENT du clic — sinon on prendrait un snapshot figé au
 * dernier render et raterait les filtres récents.
 * ============================================================================
 */

import { exportRowsToCsv, exportRowsToPdf, type ExportRow } from '../lib/exporters'

interface ExportButtonsProps {
  /** Nom de fichier (sans extension) pour le CSV, et titre pour le PDF. */
  filename: string
  /** Titre affiché en haut du PDF (souvent === libellé fonctionnel de la page). */
  pdfTitle: string
  /** Sous-titre optionnel (ex : résumé du filtre actif). */
  pdfSubtitle?: string
  /** Callback : libellés des colonnes. Évalué au clic. */
  getHeaders: () => string[]
  /** Callback : lignes de données. Évalué au clic (= filtre courant). */
  getRows: () => ExportRow[]
  /** Désactive les boutons (ex : pendant le chargement). */
  disabled?: boolean
}

/**
 * Rend deux boutons côte à côte : "Excel" (CSV téléchargé) et "PDF"
 * (fenêtre d'impression). Le style reprend les `btn-secondary` standards
 * de l'app (voir Dashboard.css / Anomalies.css).
 */
export function ExportButtons({
  filename,
  pdfTitle,
  pdfSubtitle,
  getHeaders,
  getRows,
  disabled,
}: ExportButtonsProps) {
  const handleExcel = () => {
    const headers = getHeaders()
    const rows = getRows()
    if (rows.length === 0) {
      alert('Aucune donnée à exporter avec les filtres actuels.')
      return
    }
    exportRowsToCsv(filename, headers, rows)
  }

  const handlePdf = () => {
    const headers = getHeaders()
    const rows = getRows()
    if (rows.length === 0) {
      alert('Aucune donnée à exporter avec les filtres actuels.')
      return
    }
    exportRowsToPdf(pdfTitle, headers, rows, pdfSubtitle)
  }

  return (
    <div className="export-buttons" style={{ display: 'inline-flex', gap: 8 }}>
      <button
        type="button"
        className="btn-export btn-export-excel"
        onClick={handleExcel}
        disabled={disabled}
        title="Exporter les données filtrées au format Excel (.csv)"
      >
        Excel
      </button>
      <button
        type="button"
        className="btn-export btn-export-pdf"
        onClick={handlePdf}
        disabled={disabled}
        title="Imprimer / enregistrer au format PDF"
      >
        PDF
      </button>
    </div>
  )
}
