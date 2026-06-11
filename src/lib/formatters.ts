/**
 * ============================================================================
 * FORMATEURS PARTAGÉS (nombres, dates)
 * ============================================================================
 *
 * Petits helpers pour l'affichage de valeurs numériques et de dates,
 * réutilisables dans toutes les pages.
 * ============================================================================
 */

/**
 * Formate une date SharePoint stockée comme "YYYY-MM-DDT00:00:00Z" en
 * jj/mm/aaaa locale FR, SANS décalage de fuseau horaire.
 *
 * Pourquoi ce helper plutôt que `new Date(value).toLocaleDateString('fr-FR')` ?
 * ------------------------------------------------------------------------
 * Les dates saisies via `<input type="date">` représentent un JOUR CALENDAIRE
 * (pas un instant). Le formulaire les envoie en SP sous la forme minuit UTC
 * (`form.field_0 + 'T00:00:00Z'`). Au moment de l'affichage, `new Date(...)`
 * convertit cette ISO en heure LOCALE :
 *
 *   - Browser en UTC+1/+2 (Cameroun, Paris)   → même jour affiché ✓
 *   - Browser en UTC-X                         → JOUR PRÉCÉDENT affiché ✗
 *   - Cas particulier d'heure d'été / décalage → JOUR SUIVANT possible ✗
 *
 * Pour éviter ce bug, on extrait directement les composants YYYY-MM-DD de
 * l'ISO et on construit une Date avec ces composants en LOCAL — la Date
 * obtenue représente "ce jour minuit local", donc `.toLocaleDateString()`
 * rend systématiquement le bon jour quel que soit le fuseau.
 *
 * À utiliser pour les CHAMPS DATE saisis manuellement (field_0, field_9,
 * dateAffection, dates de clôture/régularisation, etc.). Pour les VRAIS
 * timestamps avec heure significative (Created, Modified), continuer à
 * utiliser `new Date(value).toLocaleString('fr-FR')` qui donne date + heure
 * dans le fuseau local — ce qui est le comportement attendu.
 *
 * @param value String ISO ou autre, ou undefined/null
 * @param fallback Texte affiché si la date est vide ou invalide (défaut "—")
 */
export function formatDateOnlyFR(value: string | null | undefined, fallback: string = '—'): string {
  if (!value) return fallback
  // Pattern "minuit UTC" = jour calendaire
  const calendarDayMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})T00:00:00(?:\.000)?Z?$/)
  if (calendarDayMatch) {
    const year = parseInt(calendarDayMatch[1], 10)
    const month = parseInt(calendarDayMatch[2], 10) - 1
    const day = parseInt(calendarDayMatch[3], 10)
    const d = new Date(year, month, day)
    if (Number.isNaN(d.getTime())) return fallback
    return d.toLocaleDateString('fr-FR')
  }
  // Pattern "YYYY-MM-DD" pur (sans heure) → idem, on fait pareil
  const dateOnlyMatch = value.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (dateOnlyMatch) {
    const year = parseInt(dateOnlyMatch[1], 10)
    const month = parseInt(dateOnlyMatch[2], 10) - 1
    const day = parseInt(dateOnlyMatch[3], 10)
    const d = new Date(year, month, day)
    if (Number.isNaN(d.getTime())) return fallback
    return d.toLocaleDateString('fr-FR')
  }
  // Vrai timestamp (heure non-nulle) : conversion locale standard
  const d = new Date(value)
  if (Number.isNaN(d.getTime())) return fallback
  return d.toLocaleDateString('fr-FR')
}

/**
 * Formate un montant pour affichage COMPACT dans une carte de stats.
 *
 * Sur les agrégats DCPO les montants peuvent rapidement atteindre plusieurs
 * milliards de FCFA, ce qui fait déborder les cartes du dashboard. On bascule
 * sur l'ordre de grandeur du million dès qu'on dépasse 1 000 000.
 *
 * Stratégie :
 *   - valeur < 1 000 000      → affichage complet ('150 000')
 *   - valeur ≥ 1 000 000      → en millions avec max 2 décimales ('61 775,88 M')
 *
 * Pour les contextes où la précision compte (table, détail, impression),
 * continuer à utiliser `value.toLocaleString('fr-FR')` directement.
 *
 * @param value Le montant numérique à formater (toléré : NaN, Infinity → '0')
 * @returns Chaîne formatée prête à afficher
 */
export function formatMontantCompact(value: number): string {
  if (!Number.isFinite(value)) return '0'
  const abs = Math.abs(value)
  if (abs >= 1_000_000) {
    return `${(value / 1_000_000).toLocaleString('fr-FR', {
      maximumFractionDigits: 2,
    })} M`
  }
  return value.toLocaleString('fr-FR')
}
