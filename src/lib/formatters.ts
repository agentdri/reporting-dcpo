/**
 * ============================================================================
 * FORMATEURS NUMÉRIQUES PARTAGÉS
 * ============================================================================
 *
 * Petits helpers pour l'affichage de valeurs numériques dans l'UI.
 * ============================================================================
 */

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
