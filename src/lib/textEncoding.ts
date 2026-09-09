/**
 * ============================================================================
 * RÉPARATION DES CHAÎNES "MOJIBAKE" (accents mal encodés)
 * ============================================================================
 *
 * Symptôme observé : les textes accentués renvoyés par le connecteur
 * SharePoint (via @pa-client/power-code-sdk) s'affichent parfois comme
 * "immÃ©diate" au lieu de "immédiate", "dÃ©claration" au lieu de
 * "déclaration", etc.
 *
 * Cause : un octet UTF-8 multi-octets (ex: "é" = 0xC3 0xA9 en UTF-8) est
 * réinterprété comme DEUX caractères Latin-1/Windows-1252 distincts
 * (0xC3 → "Ã", 0xA9 → "©") quelque part dans le pont SDK avant d'atteindre
 * notre code. Tous les caractères accentués français (À-ÿ, U+00C0-00FF)
 * commencent par l'octet UTF-8 0xC3 → systématiquement mal réinterprétés
 * en "Ã" + un second caractère.
 *
 * fixMojibake() reconstruit la chaîne d'origine : elle réencode chaque
 * caractère en un octet Latin-1 (`escape`), puis redécode le résultat
 * comme de l'UTF-8 (`decodeURIComponent`) — l'inverse exact de la
 * corruption. Sur du texte déjà correctement encodé, cette opération lève
 * une `URIError` (séquence UTF-8 invalide) qu'on intercepte pour renvoyer
 * la chaîne inchangée : l'opération est donc sans risque même appliquée
 * "à l'aveugle" sur des données déjà saines.
 * ============================================================================
 */

/**
 * Répare une chaîne mojibake isolée. Ne fait rien (retour identique) si
 * la chaîne ne contient aucun indice de corruption ou si la réparation
 * échoue (texte déjà correctement encodé).
 */
export function fixMojibake(text: string): string {
  if (!text) return text
  // Pré-filtre : tous les caractères accentués français corrompus par ce
  // bug commencent par "Ã" (lettres accentuées, U+00C0-00FF) ou "Â"
  // (symboles/ponctuation, U+0080-00BF — ex: °, «, », non-breaking space).
  // Évite le coût try/catch sur la majorité des chaînes (ASCII pur ou
  // déjà correctement accentuées) qui n'ont pas besoin de réparation.
  if (!/[ÃÂ]/.test(text)) return text
  try {
    return decodeURIComponent(escape(text))
  } catch {
    return text
  }
}

/**
 * Applique fixMojibake récursivement à toutes les valeurs string d'un
 * objet ou tableau (profondeur illimitée). Les autres types (number,
 * boolean, null, undefined...) sont retournés inchangés.
 *
 * Utilisé pour nettoyer en un point unique les payloads renvoyés par le
 * SDK SharePoint, sans devoir toucher chaque champ individuellement dans
 * chaque page/service.
 */
export function repairMojibakeDeep<T>(value: T): T {
  if (typeof value === 'string') {
    return fixMojibake(value) as unknown as T
  }
  if (Array.isArray(value)) {
    return value.map(v => repairMojibakeDeep(v)) as unknown as T
  }
  if (value && typeof value === 'object') {
    const out: Record<string, unknown> = {}
    for (const [k, v] of Object.entries(value as Record<string, unknown>)) {
      out[k] = repairMojibakeDeep(v)
    }
    return out as T
  }
  return value
}
