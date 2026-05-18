/**
 * ============================================================================
 * SERVICE DE GESTION DES RAPPORTS QUOTIDIENS (Module 3)
 * ============================================================================
 *
 * Persistance principale : liste SharePoint DCPO_ACTIVICTE_CONTROLLER
 *
 * Mapping des champs SharePoint :
 *   ┌─────────────────────────────────┬───────────────────────────────────┐
 *   │ Colonne SharePoint              │ Champ TypeScript                  │
 *   ├─────────────────────────────────┼───────────────────────────────────┤
 *   │ ID (auto)                       │ id                                │
 *   │ Title                           │ "Rapport <date> - <contrôleur>"   │
 *   │ controlleur (Person)            │ controleurEmail + controleurName  │
 *   │ Date                            │ date (YYYY-MM-DD)                 │
 *   │ actionDeLaJournee (multi-line)  │ lines[] (sérialisées, voir plus bas)│
 *   │ nombreAnomalieDetectee          │ anomaliesDetectees                │
 *   │ observationsGlobales            │ observationsGlobales              │
 *   │ Created (auto)                  │ submittedAt                       │
 *   │ Modified (auto)                 │ updatedAt                         │
 *   └─────────────────────────────────┴───────────────────────────────────┘
 *
 * Sérialisation des lignes dans actionDeLaJournee :
 *   La liste SharePoint n'a qu'UN champ multiligne pour TOUTES les actions
 *   de la journée. On y stocke un format hybride :
 *     1. Section humaine lisible (parcourable directement dans SharePoint)
 *     2. Bloc JSON encadré par <!--LINES_JSON ... --> pour le round-trip
 *        machine fiable (parser robuste, pas de regex fragile)
 *
 * Workflow validation :
 *   La liste SharePoint actuelle ne dispose PAS des colonnes `statut` et
 *   `historiqueValidations` → ces données sont stockées en localStorage,
 *   indexées par l'ID SharePoint du rapport. Le statut par défaut au
 *   chargement est 'Soumis'.
 *
 *   Pour activer le workflow complet de validation côté SharePoint, ajouter :
 *     - statut          : Choix (Soumis / Valider / Refuser)
 *     - historiqueValidations : Plusieurs lignes de texte (JSON)
 *
 * Brouillons :
 *   Restent en localStorage (saisie en cours, pas de pollution serveur).
 * ============================================================================
 */

import { DCPO_ACTIVICTE_CONTROLLERService } from '../generated/services/DCPO_ACTIVICTE_CONTROLLERService'
import type {
  DCPO_ACTIVICTE_CONTROLLERRead,
  DCPO_ACTIVICTE_CONTROLLERWrite,
} from '../generated/models/DCPO_ACTIVICTE_CONTROLLERModel'
import { appendUrl, parseUrlList, getFileNameFromUrl } from './ticketAttachments'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE (inchangés côté API publique)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Statut de workflow d'un rapport (persisté dans la colonne SharePoint
 * `statutValidation`). Les libellés sont volontairement identiques à ceux
 * affichés dans l'UI manager (boutons "Valider" / "Invalider (Refuser)") pour
 * éviter toute divergence visuel / stockage.
 *
 * Compatibilité descendante : les anciens rapports stockés avec 'Réalisé' /
 * 'Reporté' sont automatiquement remappés à la lecture (cf. reportFromItem).
 */
export type ActivityStatus = 'Soumis' | 'Valider' | 'Refuser'

/** Une ligne d'activité (action menée par le contrôleur). */
export interface ActivityLine {
  id: string
  domaine: string
  objet: string
  action: string
  duree: number
  observation?: string
}

/**
 * Représentation enrichie d'un rapport pour le front.
 *
 * Note sur la validation manager :
 *   - `statut`     : statut courant (Soumis / Valider / Refuser) — UN SEUL
 *                    à la fois, stocké en SharePoint dans `statutValidation`
 *   - `motifRejet` : motif courant en cas de rejet — UN SEUL à la fois,
 *                    vidé quand on bascule sur Valider
 *
 *   Pas d'historique : chaque décision manager ÉCRASE la précédente.
 */
export interface ActivityReport {
  id: string
  controleurEmail: string
  controleurName: string
  date: string
  lines: ActivityLine[]
  totalHeures: number
  tempsOccupe: number
  statutJournee: 'Calme' | 'Normal' | 'Chargé' | 'Très chargé'
  statut: ActivityStatus
  motifRejet?: string
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  attachments?: { name: string; url: string }[]
  submittedAt: string
  updatedAt: string
}

/** Brouillon en cours de saisie (localStorage uniquement). */
export interface ActivityDraft {
  id?: string
  controleurEmail: string
  date: string
  lines: ActivityLine[]
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  updatedAt: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — LOCALSTORAGE (brouillons uniquement)
 *
 * Le statut + motif de validation est désormais persisté DIRECTEMENT en
 * SharePoint via les colonnes `statutValidation` et `motifRejet` ajoutées
 * à la liste DCPO_ACTIVICTE_CONTROLLER. Plus de stockage localStorage pour
 * la validation manager → données partagées entre tous les utilisateurs.
 * ────────────────────────────────────────────────────────────────────────── */

/** Préfixe pour les brouillons de saisie (par contrôleur). */
const STORAGE_KEY_DRAFT_PREFIX = 'reportingDCPO.activityDraft.v1.'

const READ_RAW = (key: string): string | null => {
  try { return window.localStorage.getItem(key) } catch { return null }
}
const WRITE_RAW = (key: string, value: string): void => {
  try { window.localStorage.setItem(key, value) } catch { /* ignore */ }
}
const REMOVE_RAW = (key: string): void => {
  try { window.localStorage.removeItem(key) } catch { /* ignore */ }
}

/** Génère un identifiant local unique (utilisé pour les ID des lignes). */
function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}


/* ══════════════════════════════════════════════════════════════════════════
 * SECTION 3 — SÉRIALISATION DES LIGNES D'ACTION DE LA JOURNÉE
 * ══════════════════════════════════════════════════════════════════════════
 *
 * CONTEXTE
 * --------
 * La liste SharePoint DCPO_ACTIVICTE_CONTROLLER ne dispose que d'UN seul
 * champ multiligne (`actionDeLaJournee`) pour stocker TOUTES les actions
 * menées par le contrôleur dans la journée. Il faut donc encoder une liste
 * structurée d'objets ActivityLine[] dans une simple string.
 *
 * CONTRAINTES
 * -----------
 *   1. Le champ doit être LISIBLE quand un manager l'ouvre directement dans
 *      l'UI SharePoint (sans passer par l'app React)
 *   2. Le champ doit être PARSEABLE de manière robuste pour reconstruire
 *      les ActivityLine[] et les afficher dans l'app
 *   3. Le format doit survivre à un éditeur SharePoint qui pourrait
 *      reformater (ajout/suppression d'espaces, normalisation des sauts de
 *      ligne)
 *
 * SOLUTION : FORMAT HYBRIDE (texte humain + JSON encadré)
 * --------------------------------------------------------
 * Le champ contient DEUX zones :
 *
 *   ZONE 1 — Bloc humain numéroté (en haut) :
 *   ┌─────────────────────────────────────────────────────┐
 *   │ 1) [Contrôle] Visite agence Bessengue — 60 min      │
 *   │    ▸ Vérification des écritures du jour             │
 *   │    Observation : RAS                                │
 *   │                                                     │
 *   │ 2) [Reporting] Rapport hebdomadaire — 90 min        │
 *   │    ▸ Compilation des chiffres                       │
 *   └─────────────────────────────────────────────────────┘
 *
 *   ZONE 2 — Bloc JSON encadré par marqueurs HTML-comment (en bas) :
 *   ┌─────────────────────────────────────────────────────┐
 *   │ <!--LINES_JSON {"v":1,"lines":[{...},{...}]}-->     │
 *   └─────────────────────────────────────────────────────┘
 *
 * RÉPARTITION DES RÔLES
 * ---------------------
 *   - La ZONE 1 sert à la LECTURE HUMAINE (managers SharePoint)
 *   - La ZONE 2 sert à la LECTURE MACHINE (round-trip app React)
 *   - À l'écriture : on génère TOUJOURS les deux ensemble
 *   - À la lecture : on essaie d'abord ZONE 2 (JSON parfait), fallback
 *     vers ZONE 1 (parsing heuristique) si JSON absent/corrompu
 *
 * VERSIONING
 * ----------
 * Le JSON contient un champ "v":1 (version du format). Si on change le
 * schéma plus tard (ex. ajout d'un champ obligatoire), on bumpe à v:2 et
 * on garde une compatibilité descendante dans parseLines().
 *
 * POURQUOI HTML-COMMENT COMME MARQUEUR ?
 * --------------------------------------
 *   - Visuellement discret (un manager qui survole le champ ne sera pas
 *     gêné par un gros pavé de JSON brut)
 *   - SharePoint multi-line text n'interprète PAS le HTML (donc le commentaire
 *     est conservé tel quel, ne devient pas vraiment "invisible")
 *   - Facile à matcher avec une regex non-greedy ([\s\S]*?)
 * ══════════════════════════════════════════════════════════════════════════ */

/**
 * Regex de capture du bloc JSON dans le champ multiligne.
 *
 * Pattern décortiqué :
 *   <!--LINES_JSON     ← marqueur d'ouverture (littéral)
 *   \s*                ← whitespace optionnel après le marqueur
 *   ([\s\S]*?)         ← GROUPE 1 = contenu JSON (lazy : s'arrête au premier -->)
 *                        [\s\S] matche TOUT (incluant newlines), à la diff de "."
 *                        qui ne matche pas \n par défaut
 *   \s*                ← whitespace optionnel avant la fermeture
 *   -->                ← marqueur de fermeture (littéral)
 *
 * Pas de flag "i" : les marqueurs sont en majuscules strictes.
 * Pas de flag "g" : on ne cherche QU'UN seul bloc (le dernier emporte si plusieurs).
 */
const JSON_BLOCK_RE = /<!--LINES_JSON\s*([\s\S]*?)\s*-->/


/**
 * SÉRIALISATION : ActivityLine[] → string multiligne SharePoint
 * =============================================================
 *
 * Cette fonction est appelée à chaque écriture vers le champ
 * `actionDeLaJournee` de DCPO_ACTIVICTE_CONTROLLER, c'est-à-dire :
 *   - createReport() lors de la soumission initiale du rapport
 *   - updateReport() lors d'une éventuelle modification ultérieure
 *
 * Étapes (vue d'ensemble) :
 *   1. Court-circuit si liste vide → string vide (l'item SP aura un champ
 *      `actionDeLaJournee` vide aussi, c'est un état valide)
 *   2. Génération de la ZONE 1 : bloc humain pour chaque ligne
 *      - Numérotation 1-indexée (UX humaine, pas 0-indexée comme un dev)
 *      - Crochets autour du domaine pour catégorisation visuelle rapide
 *      - Tiret cadratin (—) avant la durée pour démarcation propre
 *      - Indentation de 3 espaces des sous-éléments (action, observation)
 *      - Préfixe ▸ devant l'action (icône d'item-de-liste minimaliste)
 *      - Préfixe "Observation :" devant la note (texte explicite)
 *      - Skip des sous-éléments vides pour ne pas afficher "▸ " tout seul
 *   3. Génération de la ZONE 2 : JSON minifié encadré par marqueurs
 *      - Pas d'indentation (économise des bytes dans SharePoint)
 *      - Wrapper {v:1, lines:[...]} pour le versioning futur
 *   4. Concaténation : ZONE 1 + double newline + ZONE 2
 *
 * @param lines - Tableau d'ActivityLine venant du formulaire React
 * @returns String prête à être envoyée à SharePoint dans actionDeLaJournee
 */
export function serializeLines(lines: ActivityLine[]): string {
  // Court-circuit défensif : pas de lignes → champ vide
  // (évite de stocker un bloc JSON `{"v":1,"lines":[]}` inutile en SP)
  if (!lines || lines.length === 0) return ''

  // ─────────────────────────────────────────────────────────────────────
  // ZONE 1 — Construction du bloc humain
  // ─────────────────────────────────────────────────────────────────────
  // On itère sur chaque ligne avec son index (pour la numérotation 1-indexée).
  // Chaque ligne devient potentiellement 1 à 3 lignes texte :
  //   - L'en-tête (toujours)
  //   - L'action (uniquement si non vide)
  //   - L'observation (uniquement si non vide)
  // Les blocs sont ensuite joints par "\n\n" pour créer une séparation
  // visuelle claire entre activités.
  const human = lines.map((l, i) => {
    const num = i + 1                                         // numérotation 1-indexée pour l'humain
    const domaine = (l.domaine || '—').trim()                 // fallback "—" si vide (mieux qu'une chaîne vide entre crochets)
    const objet = (l.objet || '—').trim()                    // idem
    const duree = Number.isFinite(l.duree) ? l.duree : 0      // protection contre NaN/Infinity venant d'un input number cassé

    // Accumulateur des lignes texte de cette activité
    // Note : volontairement nommé "lines" en local mais shadow "lines" du
    // paramètre extérieur — pas un problème car on n'a plus besoin du paramètre
    // ici (on a déjà le `l` courant)
    const out: string[] = []

    // ─── Ligne 1 : en-tête obligatoire ────────────────────────────────
    // Format : "1) [Contrôle] Visite agence Bessengue — 60 min"
    //   - "${num})" : numérotation visible
    //   - "[${domaine}]" : crochets pour signaler la catégorie d'un coup d'œil
    //   - "${objet}" : sujet de l'action sans décoration (texte libre)
    //   - " — ${duree} min" : durée explicite avec espacement aéré
    out.push(`${num}) [${domaine}] ${objet} — ${duree} min`)

    // ─── Ligne 2 (optionnelle) : action menée ─────────────────────────
    // Indentée de 3 espaces pour signaler la subordination à l'en-tête
    // Préfixe ▸ : flèche-triangle qui rend l'item facile à parser à l'œil
    // Le `?.trim()` gère à la fois undefined ET chaîne vide
    if (l.action?.trim()) out.push(`   ▸ ${l.action.trim()}`)

    // ─── Ligne 3 (optionnelle) : observation libre ────────────────────
    // Même indentation que l'action
    // Préfixe "Observation :" en toutes lettres (vs symbole) pour différencier
    // sémantiquement de l'action (qui est CE qui a été fait, l'observation
    // étant un commentaire/note)
    if (l.observation?.trim()) out.push(`   Observation : ${l.observation.trim()}`)

    // Joint les 1-3 lignes de cette activité par un simple newline
    return out.join('\n')
  })
  // Chaque activité est ensuite séparée par DEUX newlines (ligne blanche
  // entre les blocs) pour la lisibilité.
  .join('\n\n')

  // ─────────────────────────────────────────────────────────────────────
  // ZONE 2 — Construction du bloc JSON
  // ─────────────────────────────────────────────────────────────────────
  // JSON.stringify SANS indentation pour minimiser la taille stockée.
  // Le wrapper { v: 1, lines: [...] } permet :
  //   - "v" : versionner le format pour migration future (v2 = nouveau schéma)
  //   - "lines" : préserver l'array tel-quel pour reconstruction sans perte
  //
  // ⚠ ActivityLine contient des champs string libres qui peuvent contenir des
  // caractères spéciaux (guillemets, retours à la ligne, etc.). JSON.stringify
  // les échappe automatiquement (\\n, \", etc.), donc le contenu est safe
  // dans le bloc HTML-commenté.
  const json = JSON.stringify({ v: 1, lines })

  // ─────────────────────────────────────────────────────────────────────
  // ASSEMBLAGE FINAL
  // ─────────────────────────────────────────────────────────────────────
  // Format final : "<bloc humain>\n\n<!--LINES_JSON <json>-->"
  // - Le \n\n entre les deux zones crée une séparation visuelle nette
  // - L'espace après "LINES_JSON" et avant "-->" est facultatif mais
  //   améliore la lisibilité (et la regex JSON_BLOCK_RE le tolère via \s*)
  return `${human}\n\n<!--LINES_JSON ${json}-->`
}


/**
 * DÉSÉRIALISATION : string multiligne SharePoint → ActivityLine[]
 * ===============================================================
 *
 * Cette fonction est appelée à chaque LECTURE d'un rapport depuis SharePoint
 * (listReports, getReport). Elle reconstruit la liste des actions à partir
 * du contenu textuel stocké dans `actionDeLaJournee`.
 *
 * Stratégie en 3 niveaux (du plus fiable au plus best-effort) :
 *
 *   NIVEAU 1 — Parser le bloc JSON (chemin nominal)
 *     ✅ Round-trip parfait avec serializeLines
 *     ✅ Tous les champs préservés (id, observation optionnelle, etc.)
 *     ⚠ Échoue si JSON corrompu OU bloc absent
 *
 *   NIVEAU 2 — Parsing heuristique du texte humain (fallback)
 *     ⚠ Utile pour les items créés/modifiés MANUELLEMENT dans SharePoint
 *       sans passer par l'app (ex: un manager corrige une faute via l'UI SP)
 *     ⚠ Best-effort : peut perdre l'observation, ID régénérés, etc.
 *
 *   NIVEAU 3 — Tableau vide (échec total)
 *     ✅ L'app n'explose pas, mais l'utilisateur voit un rapport sans lignes
 *
 * @param raw - Contenu brut du champ actionDeLaJournee (peut être null)
 * @returns Tableau d'ActivityLine reconstruit (ou [] si rien parsable)
 */
export function parseLines(raw: string | null | undefined): ActivityLine[] {
  // Court-circuit : null/undefined/string vide → []
  // (équivalent à un rapport "sans actions" — état valide mais inhabituel)
  if (!raw) return []

  // ─────────────────────────────────────────────────────────────────────
  // NIVEAU 1 : Tentative de parsing du bloc JSON (chemin nominal)
  // ─────────────────────────────────────────────────────────────────────
  // .match() avec une regex non-globale retourne soit null soit un objet
  // avec match[0] = match complet et match[1+] = groupes de capture.
  // Ici match[1] = contenu entre <!--LINES_JSON et -->
  const match = raw.match(JSON_BLOCK_RE)
  if (match && match[1]) {
    try {
      // Parse JSON. Cast pessimiste : on suppose que le payload pourrait
      // être de n'importe quelle forme (pas seulement notre wrapper attendu).
      const parsed = JSON.parse(match[1]) as { v?: number; lines?: ActivityLine[] }

      // Validation structurelle : on attend un objet avec `lines: array`
      if (parsed && Array.isArray(parsed.lines)) {
        // Re-mappage défensif : on reconstruit chaque ligne avec :
        //   - Génération d'un nouvel id si manquant (jamais le cas en
        //     sortie de serializeLines, mais peut arriver si quelqu'un a
        //     édité le JSON manuellement)
        //   - Coercition string sur les champs textuels (?? '' évite
        //     undefined → 'undefined' affiché dans l'UI)
        //   - Number(l.duree) || 0 : si duree est string (édition manuelle)
        //     Number() la convertit. Le || 0 fallback gère NaN/0/null.
        return parsed.lines.map(l => ({
          id: l.id ?? uid(),
          domaine: l.domaine ?? '',
          objet: l.objet ?? '',
          action: l.action ?? '',
          duree: Number(l.duree) || 0,
          observation: l.observation,
        }))
      }
      // Si on arrive ici, le JSON était valide mais pas dans le format
      // attendu (ex: array directement, ou objet sans .lines). On laisse
      // tomber dans le fallback heuristique plus bas.
    } catch (err) {
      // JSON.parse a levé : contenu corrompu (édition manuelle ratée,
      // tronquage SharePoint, etc.). On log en warn (pas error) car ce
      // n'est pas critique — le fallback va prendre le relais.
      console.warn('parseLines: bloc JSON invalide, fallback texte', err)
    }
  }

  // ─────────────────────────────────────────────────────────────────────
  // NIVEAU 2 : Fallback heuristique sur le texte humain
  // ─────────────────────────────────────────────────────────────────────
  // Format attendu : "N) [domaine] objet — duree min" suivi de ▸ action
  // et "Observation : ..." (cf. parseLinesHeuristic).
  return parseLinesHeuristic(raw)
}

/**
 * PARSER HEURISTIQUE : extrait les lignes du texte humain quand le JSON
 * est absent ou corrompu.
 *
 * Cas d'usage typique : un manager a édité l'item directement dans
 * SharePoint pour corriger une faute, et a accidentellement supprimé/cassé
 * le bloc <!--LINES_JSON ... -->. On préserve la donnée visible plutôt
 * que de retourner [] et perdre le contenu.
 *
 * Limites :
 *   - Les IDs des lignes sont REGÉNÉRÉS (originaux perdus)
 *   - L'observation est récupérée par recherche de préfixe insensible à la casse
 *   - Si le format texte a été modifié de façon non-standard, certaines
 *     lignes peuvent être ignorées silencieusement
 */
function parseLinesHeuristic(raw: string): ActivityLine[] {
  // Étape 1 : Retirer le bloc JSON s'il existe (avec contenu corrompu)
  // → on ne veut pas tenter de parser le JSON brut comme du texte humain
  // → si le bloc est absent, .replace() est un no-op (rien retiré)
  const text = raw.replace(JSON_BLOCK_RE, '').trim()
  if (!text) return []

  const out: ActivityLine[] = []

  // Étape 2 : Découpage en blocs sur les lignes blanches
  // Regex /\n\s*\n/ : un newline + optionnellement des whitespaces + un newline
  // Tolère les variations comme "\n\n", "\n \n", "\n\t\n", etc.
  const blocks = text.split(/\n\s*\n/)

  // Étape 3 : Pour chaque bloc, tenter d'extraire une ligne d'activité
  for (const block of blocks) {
    // Découper le bloc en lignes individuelles, trim chaque ligne, ignorer les vides
    const lines = block.split(/\r?\n/).map(l => l.trim()).filter(Boolean)
    if (lines.length === 0) continue

    // Étape 3.a : Parser la ligne d'EN-TÊTE
    // Pattern : "N) [domaine] objet — duree min"
    //   - ^\d+\)         : numéro suivi de )
    //   - \s*\[([^\]]+)\] : crochets, capture du domaine
    //   - \s*(.+?)        : objet (lazy pour ne pas avaler le tiret)
    //   - \s*[—-]\s*      : tiret cadratin OU tiret simple (tolérant)
    //   - (\d+)\s*min     : durée + suffixe "min"
    //   - flag /i         : case-insensitive ("MIN" tolérable)
    const headerMatch = lines[0].match(/^\d+\)\s*\[([^\]]+)\]\s*(.+?)\s*[—-]\s*(\d+)\s*min/i)
    if (!headerMatch) continue  // bloc qui n'est pas une activité → skip

    // Étape 3.b : Chercher la ligne d'action (préfixe ▸)
    // .find() retourne la première ligne matchant ; le ?. gère undefined
    // .replace(/^▸\s*/, '') retire le préfixe pour ne garder que le contenu
    const action = lines.find(l => l.startsWith('▸'))?.replace(/^▸\s*/, '') ?? ''

    // Étape 3.c : Chercher la ligne d'observation
    // toLowerCase pour matcher "Observation", "observation", "OBSERVATION"
    // /^observation\s*[:：]\s*/i : retire le préfixe "Observation :"
    //   - [:：] tolère le deux-points ASCII (:) ET le pleine-largeur (：)
    //   - flag /i : insensible à la casse
    const observation = lines.find(l => l.toLowerCase().startsWith('observation'))?.replace(/^observation\s*[:：]\s*/i, '') ?? ''

    // Étape 3.d : Construction de l'ActivityLine reconstruite
    out.push({
      id: uid(),                                  // nouvel ID (l'original est perdu)
      domaine: headerMatch[1].trim(),             // groupe 1 = domaine entre crochets
      objet: headerMatch[2].trim(),               // groupe 2 = objet
      duree: parseInt(headerMatch[3], 10) || 0,   // groupe 3 = durée. parseInt + fallback 0
      action,                                     // peut être '' si pas de ligne ▸
      observation: observation || undefined,      // undefined si vide (cohérent avec le type optionnel)
    })
  }
  return out
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — MAPPING ITEM SHAREPOINT → ACTIVITYREPORT
 * ────────────────────────────────────────────────────────────────────────── */

/** Construit un ActivityReport complet à partir d'un item SharePoint. */
function reportFromItem(item: DCPO_ACTIVICTE_CONTROLLERRead): ActivityReport {
  const lines = parseLines(item.actionDeLaJournee)
  const totals = computeTotals(lines)
  const dateStr = item.Date ? item.Date.split('T')[0] : ''

  // Pièces jointes : le champ urlPieceJointes (champ texte SP) peut contenir
  // plusieurs URLs concaténées par " | " (cf. helper appendUrl).
  // On les transforme en tableau {name, url} pour l'affichage côté UI.
  const attachmentUrls = parseUrlList(item.urlPieceJointes)
  const attachments = attachmentUrls.map(url => ({
    name: getFileNameFromUrl(url),
    url,
  }))

  // Validation manager : statut + motif lus DIRECTEMENT depuis SharePoint.
  // Un seul statut + un seul motif à la fois (pas d'historique).
  // Fallback 'Soumis' si la colonne statutValidation est vide (rapport
  // nouvellement créé, en attente de décision manager).
  //
  // Compatibilité descendante : les anciens rapports utilisaient 'Réalisé' /
  // 'Reporté' avant le renommage des libellés. On les remappe à la volée vers
  // les nouvelles valeurs (Valider / Refuser) pour éviter de perdre l'historique.
  const rawStatut = (item.statutValidation ?? '').trim()
  const statut: ActivityStatus =
    rawStatut === 'Valider' || rawStatut === 'Réalisé' ? 'Valider' :
    rawStatut === 'Refuser' || rawStatut === 'Reporté' ? 'Refuser' :
    'Soumis'

  return {
    id: String(item.ID),
    controleurEmail: item.controlleur?.Email ?? '',
    controleurName: item.controlleur?.DisplayName ?? '',
    date: dateStr,
    lines,
    totalHeures: totals.totalHeures,
    tempsOccupe: totals.tempsOccupe,
    statutJournee: totals.statutJournee,
    statut,
    motifRejet: item.motifRejet ?? undefined,
    anomaliesDetectees: item.nombreAnomalieDetectee,
    lienRapport: undefined,
    observationsGlobales: item.observationsGlobales,
    attachments,
    submittedAt: item.Created ?? '',
    updatedAt: item.Modified ?? '',
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — API PUBLIQUE : LECTURE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Liste tous les rapports depuis SharePoint, triés du plus récent au plus ancien.
 * Robuste : retourne [] en cas d'erreur (logs en console).
 */
export async function listReports(): Promise<ActivityReport[]> {
  try {
    const result = await DCPO_ACTIVICTE_CONTROLLERService.getAll({
      orderBy: ['Date desc'],
    })
    if (!result.data) return []
    return result.data.map(reportFromItem)
  } catch (err) {
    console.error('listReports error', err)
    return []
  }
}

/** Récupère un rapport par ID SharePoint. */
export async function getReport(id: string): Promise<ActivityReport | undefined> {
  try {
    const result = await DCPO_ACTIVICTE_CONTROLLERService.get(id)
    if (!result.data) return undefined
    return reportFromItem(result.data)
  } catch (err) {
    console.error('getReport error', err)
    return undefined
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — CALCULS DÉRIVÉS
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Calcule les totaux dérivés à partir des lignes :
 *   - totalHeures   : somme des durées en heures (arrondi 2 décimales)
 *   - tempsOccupe   : pourcentage d'une journée 8h occupée (cap à 100%)
 *   - statutJournee : étiquette qualitative basée sur totalHeures
 *
 * Recalculé à la lecture (pas stocké en SharePoint) → toujours cohérent
 * avec le contenu réel des lignes.
 */
export function computeTotals(lines: ActivityLine[]): { totalHeures: number; tempsOccupe: number; statutJournee: ActivityReport['statutJournee'] } {
  const totalMinutes = lines.reduce((s, l) => s + (Number.isFinite(l.duree) ? Math.max(0, l.duree) : 0), 0)
  const totalHeures = Math.round((totalMinutes / 60) * 100) / 100
  const tempsOccupe = Math.min(100, Math.round((totalMinutes / (8 * 60)) * 100))

  let statutJournee: ActivityReport['statutJournee'] = 'Normal'
  if (totalHeures < 4) statutJournee = 'Calme'
  else if (totalHeures < 7) statutJournee = 'Normal'
  else if (totalHeures < 9) statutJournee = 'Chargé'
  else statutJournee = 'Très chargé'

  return { totalHeures, tempsOccupe, statutJournee }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 7 — CRÉATION D'UN RAPPORT (SHAREPOINT)
 * ────────────────────────────────────────────────────────────────────────── */

/** Données minimales pour créer un rapport. */
export interface CreateReportInput {
  controleurEmail: string
  controleurName: string
  date: string
  lines: ActivityLine[]
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  attachments?: { name: string; url: string }[]
}

/** Format SharePoint Claims pour les champs Personne/Groupe. */
function toClaims(email: string): string {
  return `i:0#.f|membership|${email}`
}

/**
 * Crée un nouvel item SharePoint dans DCPO_ACTIVICTE_CONTROLLER.
 *
 * Étapes :
 *   1. Sérialise les lignes vers le format hybride humain+JSON
 *   2. Construit le payload SP (controlleur en format Claims)
 *   3. POST via le service généré
 *   4. Retourne le rapport reconstruit depuis l'item créé
 */
export async function createReport(input: CreateReportInput): Promise<ActivityReport> {
  const actionText = serializeLines(input.lines)
  const title = `Rapport ${input.date} - ${input.controleurName}`
  const dateISO = `${input.date}T00:00:00Z`

  const payload: Record<string, unknown> = {
    Title: title,
    Date: dateISO,
    actionDeLaJournee: actionText,
    nombreAnomalieDetectee: input.anomaliesDetectees ?? 0,
    observationsGlobales: input.observationsGlobales ?? '',
    // Statut initial : en attente de décision manager
    // (motifRejet reste vide à la création — il ne sera renseigné que sur
    //  une éventuelle décision Reporté ultérieure).
    statutValidation: 'Soumis',
    motifRejet: '',
  }
  if (input.controleurEmail) {
    payload.controlleur = {
      '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
      Claims: toClaims(input.controleurEmail),
    }
  }

  const result = await DCPO_ACTIVICTE_CONTROLLERService.create(
    payload as Omit<DCPO_ACTIVICTE_CONTROLLERWrite, 'ID'>,
  )
  if (!result.success || !result.data) {
    throw new Error(result.error?.message ?? 'Échec de la création du rapport.')
  }
  return reportFromItem(result.data)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 8 — MISE À JOUR D'UN RAPPORT
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Patch un rapport existant en SharePoint.
 *
 * Limitation : les calculs dérivés (totalHeures, tempsOccupe, statutJournee)
 * sont recalculés à la lecture, donc pas besoin de les mettre à jour ici.
 * Si patch.lines est fourni, on ré-encode via serializeLines().
 */
export async function updateReport(id: string, patch: Partial<ActivityReport>): Promise<ActivityReport | undefined> {
  const payload: Record<string, unknown> = {}
  if (patch.lines) payload.actionDeLaJournee = serializeLines(patch.lines)
  if (patch.anomaliesDetectees !== undefined) payload.nombreAnomalieDetectee = patch.anomaliesDetectees
  if (patch.observationsGlobales !== undefined) payload.observationsGlobales = patch.observationsGlobales
  if (patch.date) payload.Date = `${patch.date}T00:00:00Z`

  if (Object.keys(payload).length > 0) {
    try {
      await DCPO_ACTIVICTE_CONTROLLERService.update(id, payload as never)
    } catch (err) {
      console.error('updateReport error', err)
      return undefined
    }
  }
  return getReport(id)
}

/**
 * Ajoute (concatène) une ou plusieurs URLs de pièces jointes au champ
 * `urlPieceJointes` du rapport SharePoint, sans écraser celles existantes.
 *
 * Pattern identique à celui utilisé pour les anomalies (cf. appendUrl du lib
 * ticketAttachments.ts) : les URLs sont concaténées par " | " dans le champ
 * texte. Le helper appendUrl gère le dédoublonnage et la mise en forme.
 *
 * Étapes :
 *   1. Lit l'item courant pour récupérer la valeur existante de urlPieceJointes
 *   2. Pour chaque URL nouvelle, applique appendUrl (qui déduplique)
 *   3. Met à jour SharePoint avec la chaîne concaténée
 *
 * Utilisé après l'upload de pièces jointes via uploadActivityAttachment :
 * le workflow Power Automate retourne l'URL de chaque fichier ajouté, et on
 * persiste ces URLs dans le champ pour pouvoir les afficher côté lecture.
 *
 * @param reportId - ID SharePoint du rapport cible
 * @param newUrls  - URLs à ajouter au champ (typiquement les URLs retournées
 *                   par les uploads Power Automate, dans l'ordre des fichiers)
 */
export async function appendReportAttachmentUrls(reportId: string, newUrls: string[]): Promise<void> {
  // Court-circuit : aucune URL à ajouter → no-op (évite un round-trip inutile)
  const cleanUrls = newUrls.filter(u => !!u && u.trim().length > 0)
  if (cleanUrls.length === 0) return

  // 1. Lire la valeur actuelle pour ne pas écraser les URLs déjà stockées
  let existing = ''
  try {
    const result = await DCPO_ACTIVICTE_CONTROLLERService.get(reportId)
    existing = result.data?.urlPieceJointes ?? ''
  } catch (err) {
    console.error('appendReportAttachmentUrls: échec lecture item', err)
    // On continue avec une chaîne vide — pire cas on perd les anciennes URLs,
    // mais c'est mieux que de ne pas écrire les nouvelles.
  }

  // 2. Concaténer chaque nouvelle URL via appendUrl
  //    appendUrl gère :
  //      - le dédoublonnage (si une URL est déjà présente, elle n'est pas réajoutée)
  //      - le format avec séparateur " | "
  //      - la propreté finale (trim, etc.)
  const concatenated = cleanUrls.reduce(
    (acc, url) => appendUrl(acc, url),
    existing,
  )

  // 3. Persister la nouvelle valeur dans SharePoint
  try {
    await DCPO_ACTIVICTE_CONTROLLERService.update(reportId, {
      urlPieceJointes: concatenated,
    } as Partial<Omit<DCPO_ACTIVICTE_CONTROLLERWrite, 'ID'>>)
  } catch (err) {
    console.error('appendReportAttachmentUrls: échec update item', err)
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 9 — VALIDATION MANAGER (persistance SharePoint)
 *
 * Le statut et le motif de rejet sont stockés DIRECTEMENT dans la liste
 * SharePoint DCPO_ACTIVICTE_CONTROLLER via les colonnes :
 *   - statutValidation : 'Soumis' / 'Valider' / 'Refuser'
 *   - motifRejet       : texte libre, pertinent uniquement si Refuser
 *
 * Règle métier : UN seul statut + UN seul motif à la fois (pas d'historique).
 * Chaque décision manager écrase la précédente, et basculer sur Valider vide
 * automatiquement le motifRejet pour la cohérence.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Entrée d'une validation manager.
 *
 * Les champs `manager` et `managerEmail` sont conservés dans l'interface pour
 * compatibilité avec l'UI existante, mais ils NE SONT PAS persistés en
 * SharePoint (l'historique n'est plus stocké — seul le statut courant l'est).
 * L'identité du dernier validateur peut être retrouvée via les méta SharePoint
 * `Editor` / `Modified` du record.
 */
export interface ValidationInput {
  manager: string
  managerEmail?: string
  decision: 'Valider' | 'Refuser'
  motif?: string
}

/**
 * Applique une décision manager au rapport SharePoint.
 *
 * Comportement :
 *   - decision = 'Valider' → statutValidation = 'Valider', motifRejet vidé
 *   - decision = 'Refuser' → statutValidation = 'Refuser', motifRejet = motif
 *
 * Retourne le rapport rechargé (avec le nouveau statut + motif).
 */
export async function validateReport(id: string, input: ValidationInput): Promise<ActivityReport | undefined> {
  // Préparation du payload : on écrit TOUJOURS les deux colonnes en même temps
  // pour garantir la cohérence (un Valider ne doit pas garder un ancien motif).
  const payload: Record<string, unknown> = {
    statutValidation: input.decision,
    motifRejet: input.decision === 'Refuser' ? (input.motif?.trim() ?? '') : '',
  }
  try {
    await DCPO_ACTIVICTE_CONTROLLERService.update(
      id,
      payload as Partial<Omit<DCPO_ACTIVICTE_CONTROLLERWrite, 'ID'>>,
    )
  } catch (err) {
    console.error('validateReport: échec update SharePoint', err)
    return undefined
  }
  return getReport(id)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 10 — BROUILLONS (localStorage)
 * ────────────────────────────────────────────────────────────────────────── */

export function loadDraft(controleurEmail: string): ActivityDraft | undefined {
  const raw = READ_RAW(STORAGE_KEY_DRAFT_PREFIX + controleurEmail.toLowerCase())
  if (!raw) return undefined
  try {
    return JSON.parse(raw) as ActivityDraft
  } catch {
    return undefined
  }
}

export function saveDraft(draft: ActivityDraft): void {
  WRITE_RAW(STORAGE_KEY_DRAFT_PREFIX + draft.controleurEmail.toLowerCase(), JSON.stringify({
    ...draft,
    updatedAt: new Date().toISOString(),
  }))
}

export function clearDraft(controleurEmail: string): void {
  REMOVE_RAW(STORAGE_KEY_DRAFT_PREFIX + controleurEmail.toLowerCase())
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 11 — HELPERS UI (inchangés)
 * ────────────────────────────────────────────────────────────────────────── */

export function makeEmptyLine(): ActivityLine {
  return { id: uid(), domaine: '', objet: '', action: '', duree: 30, observation: '' }
}

export const ACTIVITY_DOMAINES = [
  'Contrôle',
  'Reporting',
  'Investigation',
  'Réunion',
  'Formation',
  'Audit',
  'Suivi anomalie',
  'Administratif',
  'Autre',
]

export const ACTIVITY_LINE_LIMITS = {
  actionMinLength: 3,
  actionMaxLength: 500,
  dureeMin: 5,
  dureeMax: 600,
}

export interface LineValidationError {
  index: number
  field: keyof ActivityLine
  message: string
}

export function validateLines(lines: ActivityLine[]): LineValidationError[] {
  const errors: LineValidationError[] = []
  if (lines.length === 0) {
    errors.push({ index: -1, field: 'action', message: 'Au moins une ligne est requise.' })
    return errors
  }
  lines.forEach((line, index) => {
    if (!line.domaine) errors.push({ index, field: 'domaine', message: 'Domaine requis.' })
    if (!line.objet?.trim()) errors.push({ index, field: 'objet', message: 'Objet requis.' })
    const actionLen = line.action?.trim().length ?? 0
    if (actionLen < ACTIVITY_LINE_LIMITS.actionMinLength) {
      errors.push({ index, field: 'action', message: `Action: ${ACTIVITY_LINE_LIMITS.actionMinLength} caractères minimum.` })
    } else if (actionLen > ACTIVITY_LINE_LIMITS.actionMaxLength) {
      errors.push({ index, field: 'action', message: `Action: ${ACTIVITY_LINE_LIMITS.actionMaxLength} caractères maximum.` })
    }
    const duree = Number(line.duree)
    if (!Number.isFinite(duree) || duree < ACTIVITY_LINE_LIMITS.dureeMin) {
      errors.push({ index, field: 'duree', message: `Durée: ${ACTIVITY_LINE_LIMITS.dureeMin} min minimum.` })
    } else if (duree > ACTIVITY_LINE_LIMITS.dureeMax) {
      errors.push({ index, field: 'duree', message: `Durée: ${ACTIVITY_LINE_LIMITS.dureeMax} min maximum.` })
    }
  })
  return errors
}
