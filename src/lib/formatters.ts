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
 * Échelle d'unités compactes (suffixes SI + usage financier).
 *
 *   M  = million        (10^6)
 *   Md = milliard        (10^9)  — équivalent SI : G (giga)
 *   T  = tera            (10^12) — équivalent FR : billion
 *   P  = peta            (10^15)
 *   E  = exa             (10^18)
 *   Z  = zetta           (10^21)
 *   Y  = yotta           (10^24)
 *   R  = ronna           (10^27)
 *   Q  = quetta          (10^30)
 *
 * Choix de "Md" pour le milliard (au lieu de "G") : pratique courante dans
 * les contextes financiers francophones (rapports BEAC, bilans bancaires…).
 * Au-delà de quetta on bascule en notation scientifique car (1) les suffixes
 * sont inconnus du grand public et (2) le contexte DCPO ne rencontrera jamais
 * de tels montants.
 *
 * Tableau ordonné DECROISSANT (du plus grand au plus petit seuil) — la boucle
 * dans formatMontantCompact prend la première unité dont le seuil est atteint.
 */
const COMPACT_UNITS: ReadonlyArray<{ threshold: number; suffix: string }> = [
  { threshold: 1e30, suffix: 'Q' },
  { threshold: 1e27, suffix: 'R' },
  { threshold: 1e24, suffix: 'Y' },
  { threshold: 1e21, suffix: 'Z' },
  { threshold: 1e18, suffix: 'E' },
  { threshold: 1e15, suffix: 'P' },
  { threshold: 1e12, suffix: 'T' },
  { threshold: 1e9, suffix: 'Md' },
  { threshold: 1e6, suffix: 'M' },
]

/**
 * Formate un montant pour affichage COMPACT dans une carte de stats KPI.
 *
 * Stratégie :
 *   - valeur < 1 000 000        → affichage complet avec séparateurs FR
 *                                  ('999 999') — lisibilité maximale pour
 *                                  les montants "normaux" (sous le million)
 *   - 1 M ≤ valeur < 10^33      → abréviation SI/financière, décimales
 *                                  adaptatives :
 *                                    • |scaled| < 10  → 2 décimales ('1,25 Md')
 *                                    • |scaled| < 100 → 1 décimale  ('12,5 Md')
 *                                    • sinon           → 0 décimale  ('125 Md')
 *                                  Une carte affiche donc au pire 3 chiffres
 *                                  significatifs + suffixe (~6-7 caractères),
 *                                  ce qui rentre dans la grille de cartes.
 *   - valeur ≥ 10^33            → notation scientifique compacte
 *                                  ('1,23×10^33') — fallback de sécurité
 *
 * Couverture pratique :
 *   - Précision exacte JavaScript : jusqu'à 2^53 − 1 ≈ 9 × 10^15 (P / peta).
 *     Au-delà, les derniers chiffres significatifs peuvent être imprécis mais
 *     l'ordre de grandeur affiché reste correct (suffisant pour un KPI).
 *   - Limite numérique JS : Number.MAX_VALUE ≈ 1,8 × 10^308 — on peut donc
 *     toujours afficher quelque chose, même pour des valeurs absurdes.
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
  // Sous le million : affichage complet (lisibilité prioritaire)
  if (abs < 1_000_000) return value.toLocaleString('fr-FR')
  // Au-delà de quetta (10^30) : notation scientifique de secours
  if (abs >= 1e33) {
    const exp = Math.floor(Math.log10(abs))
    const mantissa = value / Math.pow(10, exp)
    return `${mantissa.toLocaleString('fr-FR', { maximumFractionDigits: 2 })}×10^${exp}`
  }
  // Cas standard : on cherche le plus grand suffixe applicable
  for (const { threshold, suffix } of COMPACT_UNITS) {
    if (abs >= threshold) {
      const scaled = value / threshold
      const absScaled = Math.abs(scaled)
      const maxFrac = absScaled < 10 ? 2 : absScaled < 100 ? 1 : 0
      return `${scaled.toLocaleString('fr-FR', {
        maximumFractionDigits: maxFrac,
      })} ${suffix}`
    }
  }
  // Filet de sécurité (théoriquement inatteignable car le test < 1e6 est
  // déjà fait au-dessus).
  return value.toLocaleString('fr-FR')
}
