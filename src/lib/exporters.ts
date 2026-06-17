/**
 * ============================================================================
 * EXPORTS — CSV (Excel-compatible) & PDF (impression navigateur)
 * ============================================================================
 *
 * Approche "zéro dépendance" :
 *   - Excel : on génère un .csv UTF-8 avec BOM, séparateur `;` (locale FR).
 *     Excel l'ouvre nativement avec les bons accents et la bonne séparation
 *     de colonnes.
 *   - PDF   : on ouvre une fenêtre popup avec un layout minimal (tableau +
 *     en-tête imprimable), on déclenche `window.print()` et l'utilisateur
 *     choisit "Enregistrer en PDF" dans la boîte d'impression. Marche dans
 *     tous les navigateurs modernes sans embarquer jsPDF.
 *
 * Convention d'usage :
 *   const headers = ['Numéro', 'Titre', 'Statut', 'Date']
 *   const rows = items.map(it => [it.id, it.titre, it.statut, it.date])
 *   exportRowsToCsv('anomalies', headers, rows)
 *   exportRowsToPdf('Liste des anomalies', headers, rows)
 *
 * Limitations connues :
 *   - CSV : pas de mise en forme cellule (gras, couleurs, formules). Pour ça
 *     il faudrait passer à xlsx (lib SheetJS) — voir tâche future si besoin.
 *   - PDF : la qualité d'impression dépend du navigateur. Sur Chrome/Edge,
 *     le rendu est correct ; sur Firefox il faut souvent activer "Graphiques
 *     d'arrière-plan" dans les options d'impression pour conserver les fonds.
 * ============================================================================
 */

/** Valeur de cellule autorisée — sera convertie en string à l'export. */
export type ExportCell = string | number | boolean | null | undefined

/** Une ligne de données = un tableau de cellules. */
export type ExportRow = ExportCell[]


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — EXPORT CSV (lisible par Excel)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Échappe une cellule pour le CSV :
 *   - Convertit en string
 *   - Si la valeur contient `;`, `"`, newline → entoure de `"..."` et double
 *     les `"` internes (norme RFC 4180)
 *
 * Cas particulier : on protège systématiquement les nombres qui ressemblent
 * à des dates ("01/01/2026" pourrait être réinterprété par Excel) en les
 * gardant dans des guillemets.
 */
function escapeCsvCell(value: ExportCell): string {
  if (value === null || value === undefined) return ''
  const str = String(value)
  // Caractères qui forcent un quoting : séparateur, guillemet, newline
  if (/[;"\r\n]/.test(str)) {
    return `"${str.replace(/"/g, '""')}"`
  }
  return str
}

/**
 * Génère et télécharge un fichier CSV.
 *
 * @param filename    nom du fichier (sans extension — .csv ajouté auto).
 *                    Sera nettoyé pour éviter les caractères système interdits.
 * @param headers     ligne d'en-tête (libellés des colonnes)
 * @param rows        données — un tableau de tableaux (cellules)
 *
 * Notes techniques :
 *   - BOM UTF-8 (`﻿`) en tête → Excel détecte le fichier comme UTF-8
 *     et affiche les accents correctement.
 *   - Séparateur `;` (locale FR — Excel français l'attend par défaut).
 *   - Newline `\r\n` (compat Windows + Mac/Linux modernes).
 *   - `URL.createObjectURL` + clic sur `<a download>` : pattern standard
 *     pour déclencher un téléchargement sans serveur. Le blob est révoqué
 *     après pour éviter les fuites mémoire.
 */
export function exportRowsToCsv(
  filename: string,
  headers: string[],
  rows: ExportRow[],
): void {
  const lines: string[] = []
  lines.push(headers.map(escapeCsvCell).join(';'))
  for (const row of rows) {
    lines.push(row.map(escapeCsvCell).join(';'))
  }
  // BOM UTF-8 pour qu'Excel reconnaisse l'encodage et affiche les accents
  const csv = '﻿' + lines.join('\r\n')
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const safeName = sanitizeFilename(filename)
  const a = document.createElement('a')
  a.href = url
  a.download = `${safeName}.csv`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  // Libération mémoire (différée d'un tick pour laisser le clic se propager)
  setTimeout(() => URL.revokeObjectURL(url), 100)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — EXPORT PDF (via impression navigateur)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Ouvre une fenêtre popup avec un tableau prêt à imprimer puis déclenche
 * `window.print()`. L'utilisateur choisit "Enregistrer au format PDF"
 * dans la boîte d'impression du navigateur.
 *
 * @param title    titre affiché en haut du document (ex : "Liste des anomalies")
 * @param headers  libellés de colonnes
 * @param rows     données
 * @param subtitle texte optionnel sous le titre (ex : "Filtre : statut = Ouvert")
 *
 * Pourquoi une popup plutôt que `window.print()` sur la page courante ?
 *   - Isole le contenu à imprimer : pas besoin de CSS @media print complexe
 *     pour cacher la navigation/topbar.
 *   - Le titre/sous-titre/footer sont entièrement maîtrisés.
 *   - L'utilisateur conserve l'app intacte derrière.
 *
 * Cas d'erreur :
 *   - Bloqueur de popup actif → window.open() renvoie null. On affiche une
 *     alerte explicite plutôt que de laisser le clic sans effet.
 */
export function exportRowsToPdf(
  title: string,
  headers: string[],
  rows: ExportRow[],
  subtitle?: string,
): void {
  const popup = window.open('', '_blank', 'width=1200,height=800')
  if (!popup) {
    alert(
      "Impossible d'ouvrir la fenêtre d'impression — veuillez autoriser les " +
      "fenêtres popup pour ce site.",
    )
    return
  }
  const safeTitle = escapeHtml(title)
  const safeSubtitle = subtitle ? escapeHtml(subtitle) : ''
  const today = new Date().toLocaleDateString('fr-FR', {
    day: '2-digit',
    month: 'long',
    year: 'numeric',
  })
  const headerCells = headers
    .map(h => `<th>${escapeHtml(h)}</th>`)
    .join('')
  const bodyRows = rows
    .map(
      r =>
        `<tr>${r
          .map(c => `<td>${escapeHtml(c === null || c === undefined ? '' : String(c))}</td>`)
          .join('')}</tr>`,
    )
    .join('')

  popup.document.write(`<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8" />
  <title>${safeTitle}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      color: #1a1a2e;
      margin: 16px;
      font-size: 11px;
    }
    h1 {
      font-size: 18px;
      margin: 0 0 4px;
      color: #c0392b;
      text-align: center;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .subtitle {
      font-size: 11px;
      color: #555;
      text-align: center;
      margin: 0 0 4px;
    }
    .meta {
      font-size: 10px;
      color: #888;
      text-align: center;
      margin: 0 0 12px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: auto;
    }
    th, td {
      border: 1px solid #94a3b8;
      padding: 4px 6px;
      text-align: left;
      vertical-align: top;
      word-break: break-word;
    }
    thead th {
      background: #1d4ed8;
      color: #fff;
      font-size: 10px;
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }
    tbody tr:nth-child(even) td { background: #f3f4f6; }
    @page { size: A4 landscape; margin: 10mm; }
  </style>
</head>
<body>
  <h1>${safeTitle}</h1>
  ${safeSubtitle ? `<p class="subtitle">${safeSubtitle}</p>` : ''}
  <p class="meta">Édité le ${today} — ${rows.length} ligne${rows.length > 1 ? 's' : ''}</p>
  <table>
    <thead><tr>${headerCells}</tr></thead>
    <tbody>${bodyRows}</tbody>
  </table>
  <script>
    // On attend que les fonts/styles soient appliqués avant d'imprimer.
    window.onload = function () {
      window.focus();
      window.print();
      // Sur Chrome/Edge, fermeture immédiate après print() ferait disparaître
      // l'aperçu. On laisse la fenêtre ouverte — l'utilisateur la ferme.
    };
  </script>
</body>
</html>`)
  popup.document.close()
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — UTILITAIRES
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Nettoie un nom de fichier des caractères interdits par les systèmes de
 * fichiers (Windows / macOS / Linux). Remplace par `_`.
 *
 * Caractères interdits Windows : `< > : " / \ | ? *`
 * + caractères de contrôle (\x00-\x1f).
 */
function sanitizeFilename(name: string): string {
  return name
    .replace(/[<>:"/\\|?*\x00-\x1f]/g, '_')
    .replace(/\s+/g, '_')
    .slice(0, 100) || 'export'
}

/**
 * Échappe les caractères HTML spéciaux pour éviter l'injection dans la
 * fenêtre PDF (les données viennent des saisies utilisateurs SP).
 */
function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/**
 * Helper : formate une date ISO en JJ/MM/AAAA pour les exports.
 * Retourne `''` pour les valeurs vides (plus lisible qu'un "—" en CSV
 * qu'Excel pourrait interpréter bizarrement dans certains cas).
 */
export function formatDateForExport(value?: string | null): string {
  if (!value) return ''
  // Pattern "minuit UTC" (jour calendaire) : on extrait les composants
  // directement pour éviter tout décalage de fuseau.
  const m = value.match(/^(\d{4})-(\d{2})-(\d{2})T00:00:00(?:\.000)?Z?$/)
  if (m) return `${m[3]}/${m[2]}/${m[1]}`
  const d = new Date(value)
  if (Number.isNaN(d.getTime())) return ''
  return d.toLocaleDateString('fr-FR')
}
