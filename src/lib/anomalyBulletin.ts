/**
 * ============================================================================
 * MAPPING ANOMALIE → BULLETIN CONSOLIDÉ (Module 2)
 * ============================================================================
 *
 * Rôle :
 *   Convertit un ticket SharePoint brut (DCPO_LISTE_ANORMALIE) en un objet
 *   "bulletin" prêt à afficher : champs résolus (libellés au lieu d'IDs),
 *   dates parsées, timeline construite, filtres applicables.
 *
 * Pourquoi cette couche ?
 *   - Évite de dupliquer la logique de formatage dans le JSX
 *   - Permet de centraliser les fallbacks (ex : description = field_4 sinon
 *     field_3 sinon Title)
 *   - Encapsule la timeline (cycle de vie) reconstituée à partir des dates
 *     éparpillées sur l'item SharePoint
 *
 * Utilisé par :
 *   - AnomalyBulletins.tsx (Module 2) — vue principale
 *   - À terme aussi exploitable depuis Anomalies.tsx pour des affichages riches
 * ============================================================================
 */

import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import { findAgenceLabel, findReseauLabel } from './spReferenceRows'


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DE FILTRAGE ET DE STATUT
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Statut de filtrage des bulletins.
 *   - 'Tous'   : pas de filtre statut
 *   - 'Resolu' : ne montrer que les anomalies résolues
 *   - 'Clos'   : ne montrer que les anomalies closes
 *
 * Note : un bulletin n'existe QUE pour les anomalies Resolu/Clos
 * (cf. isBulletinTicket). Les Ouvert/En cours sont filtrés en amont.
 */
export type BulletinStatus = 'Resolu' | 'Clos' | 'Tous'

/**
 * Critères de filtrage côté client appliqués à la liste des bulletins.
 * Tous les champs sont des strings (pratique pour binder directement
 * sur les inputs HTML).
 */
export interface BulletinFilters {
  search: string             // recherche libre dans tous les textes du bulletin
  status: BulletinStatus
  classification: string     // ex : 'Operationnel', 'Fraude', 'Commercial'
  criticite: string          // ex : 'Faible', 'Moyenne', 'Haute', 'Critique'
  agence: string             // ID de l'agence (string pour cohérence avec le select)
  reseau: string             // ID du réseau (string)
  domaineActivite: string    // ex : 'Crédit', 'Caisse', 'Conformité'...
  agent: string              // nom OU email de l'auteur (recherche partielle)
  affecte: string            // nom OU email de la personne affectée
  declarant: string          // nom OU email du déclarant (recherche partielle)
  dateFrom: string           // date min de clôture (format YYYY-MM-DD)
  dateTo: string             // date max de clôture
}

/** État initial des filtres : tout vide, statut "Tous". */
export const EMPTY_BULLETIN_FILTERS: BulletinFilters = {
  search: '',
  status: 'Tous',
  classification: '',
  criticite: '',
  agence: '',
  reseau: '',
  domaineActivite: '',
  agent: '',
  affecte: '',
  declarant: '',
  dateFrom: '',
  dateTo: '',
}

/**
 * Valeurs de field_10 (statut SharePoint) qui qualifient un ticket
 * comme "bulletin" (i.e. accessible depuis Module 2).
 *
 * Utilisé en double :
 *   1. Côté serveur dans la requête $filter (`field_10 eq 'Resolu' or ...`)
 *   2. Côté client dans isBulletinTicket() pour re-filtrer après fetch
 */
export const BULLETIN_STATUS_VALUES = ['Resolu', 'Clos']


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — TIMELINE (CYCLE DE VIE D'UN TICKET)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Une étape unique sur la timeline d'un bulletin.
 *
 *   - date        : ISO timestamp (string), peut être absent pour notes libres
 *   - label       : titre court de l'étape (ex : "Déclaration", "Clôture")
 *   - description : précision optionnelle (ex : nom du déclarant)
 *   - type        : catégorie pour la stylisation visuelle (couleur du dot,
 *                   icône, etc.) — voir AnomalyBulletins.css `.timeline-*`
 */
export interface LifecycleStep {
  date?: string
  label: string
  description?: string
  type: 'declaration' | 'survenance' | 'opening' | 'regularization' | 'closure' | 'log' | 'event'
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — STRUCTURE DU BULLETIN CONSOLIDÉ
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Bulletin enrichi prêt à être affiché. C'est la forme cible produite
 * par buildConsolidatedBulletin() à partir d'un ticket brut.
 *
 * Champs majeurs :
 *   - ticket            : référence à l'objet brut (utile pour drill-down
 *                         vers les pièces jointes natives, urlPieceJointe…)
 *   - numero / titre    : identifiants visuels affichés en gros (ex "T-44")
 *   - *Label            : libellés résolus depuis les référentiels
 *   - *Name             : noms des personnes impliquées (DisplayName)
 *   - *Date             : dates parsées en objets Date (undefined si absent)
 *   - delayDays         : délai en jours entre déclaration et clôture
 *   - description, cause*, observations : textes nettoyés (HTML stripped)
 *   - actions           : liste des actions menées, chaque ligne séparée
 *   - lifecycle         : timeline triée chronologiquement
 *   - sharePointUrl     : lien direct vers l'item dans l'UI SharePoint
 */
export interface ConsolidatedBulletin {
  ticket: DCPO_LISTE_ANORMALIERead
  numero: string
  titre: string
  agenceLabel: string
  reseauLabel: string
  declarantName: string
  auteurName: string
  affecteName: string
  declarationDate?: Date
  openingDate?: Date
  regularizationDate?: Date
  closureDate?: Date
  delayDays?: number
  classification: string
  criticite: string
  statut: string
  montant?: number
  description: string
  causeImmediate: string
  causeRacine: string
  observations: string
  actions: string[]
  domaine: string
  natureRisque: string
  /**
   * Mode de traitement appliqué à la clôture de l'anomalie
   * (champ SP `typeSanction`). Vide tant que l'anomalie n'a pas été
   * clôturée — c'est le contrôleur qui le renseigne dans le formulaire
   * de résolution.
   */
  typeSanction: string
  occurrences?: number
  lifecycle: LifecycleStep[]
  sharePointUrl?: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — HELPERS BAS-NIVEAU (NETTOYAGE / PARSING)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Nettoie un texte HTML pour ne garder que le contenu textuel.
 *
 * Pourquoi ? SharePoint stocke souvent les textes longs (cause, description)
 * en HTML, mais on veut afficher du texte plain dans les cards/timeline.
 *
 * Méthode : utiliser DOMParser du navigateur pour parser proprement
 * (gère les entités HTML, les tags imbriqués, etc.) plutôt qu'une regex
 * fragile.
 */
function stripHtml(html?: string | null): string {
  if (!html) return ''
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
}

/**
 * Parse une string ISO en Date, mais retourne undefined plutôt qu'une
 * Date invalide.
 *
 * Use case : SharePoint peut renvoyer null, undefined, '', ou des strings
 * non parseables ("0001-01-01T00:00:00Z"). On veut tout normaliser sur
 * Date | undefined pour pouvoir tester `if (date)` côté UI.
 *
 * /!\ GESTION DU FUSEAU HORAIRE
 * ----------------------------
 * Les dates SAISIES via `<input type="date">` représentent un JOUR CALENDAIRE
 * (pas un instant précis). Le formulaire les stocke en SP comme
 * `"YYYY-MM-DDT00:00:00Z"` (minuit UTC). Si on fait `new Date(...)` puis
 * `.toLocaleDateString('fr-FR')`, le résultat dépend du fuseau du navigateur :
 *   - UTC+1 (WAT/CET) → "DD/MM" (correct)
 *   - UTC+2 (CEST)    → "DD/MM" (correct)
 *   - UTC-3 (Brésil)  → "DD-1/MM" (jour précédent — bug!)
 *   - UTC+24h hypothétique → "DD+1/MM" (jour suivant — bug!)
 *
 * Pour ÉVITER ce décalage, quand on détecte le pattern "minuit UTC"
 * (= jour calendaire), on construit la Date avec les composants LOCAUX
 * (`new Date(year, month, day)`) → la Date représente "ce jour minuit local"
 * et `.toLocaleDateString()` rend systématiquement le bon jour, quel que
 * soit le fuseau de l'utilisateur.
 *
 * Pour les VRAIS instants (créés avec `new Date().toISOString()` — qui ont
 * une heure non-nulle), on garde le comportement standard : conversion en
 * heure locale (logique car c'est bien un moment dans le temps).
 */
function parseDateSafe(value?: string | null): Date | undefined {
  if (!value) return undefined
  // Pattern "minuit UTC" = date calendrier (saisie via input date).
  // Tolère les variations ".000" sur les millisecondes et l'absence du "Z"
  // (qui n'arrive normalement pas mais reste défensif).
  const calendarDayMatch = value.match(
    /^(\d{4})-(\d{2})-(\d{2})T00:00:00(?:\.000)?Z?$/,
  )
  if (calendarDayMatch) {
    const year = parseInt(calendarDayMatch[1], 10)
    const month = parseInt(calendarDayMatch[2], 10) - 1
    const day = parseInt(calendarDayMatch[3], 10)
    const d = new Date(year, month, day)
    return Number.isNaN(d.getTime()) ? undefined : d
  }
  // Vrai timestamp (ex: dateOuvertureTicket, Created, Modified)
  // → conversion en heure locale normale.
  const d = new Date(value)
  return Number.isNaN(d.getTime()) ? undefined : d
}

/**
 * Calcule la différence en jours entiers entre deux dates.
 *
 *   - 1000 * 60 * 60 * 24 = nombre de millisecondes dans une journée
 *   - Math.round = arrondi à l'entier le plus proche (au lieu de tronquer)
 *   - Math.max(0, …) : garantit qu'on ne renvoie jamais un délai négatif
 *     (cas où end < start, donnée corrompue)
 *
 * Retourne undefined si l'une des deux dates est manquante.
 */
function diffInDays(start?: Date, end?: Date): number | undefined {
  if (!start || !end) return undefined
  const ms = end.getTime() - start.getTime()
  if (Number.isNaN(ms)) return undefined
  return Math.max(0, Math.round(ms / (1000 * 60 * 60 * 24)))
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — TYPE ÉTENDU & PICKERS DE FALLBACK
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Extension du type SharePoint généré pour inclure des champs MÉTIER
 * supplémentaires qui peuvent exister sur la liste mais ne sont pas
 * dans le modèle généré officiel (DCPO_LISTE_ANORMALIEModel).
 *
 * Pourquoi ?
 *   La liste SharePoint peut avoir évolué (ajout de field_11..23, ou de
 *   colonnes nommées comme causeImmediate, natureRisque…). Le modèle généré
 *   ne reflète qu'un sous-ensemble. Plutôt que de modifier le code généré
 *   (qui serait écrasé), on cast localement pour accéder à ces champs
 *   "additionnels" sans erreur TypeScript.
 *
 * Si un de ces champs n'existe pas réellement sur la liste, l'accès renvoie
 * simplement undefined → le code fallback prend le relais.
 */
interface ExtendedTicket extends DCPO_LISTE_ANORMALIERead {
  field_11?: string
  field_12?: string
  field_13?: string
  field_14?: string
  field_15?: string
  field_16?: string
  field_17?: string
  field_18?: string
  field_19?: string
  field_20?: number
  field_21?: string
  field_22?: string
  field_23?: string  // journal d'événements (texte multi-lignes)
  natureRisque?: string
  domaineAnomalie?: string
  causeImmediate?: string
  causeRacine?: string
  actionsAMener?: string
  observationsBulletin?: string
}

/**
 * Cherche la première chaîne non-vide dans une liste de candidats.
 *
 * Use case : robustesse de l'affichage. Le bulletin tente plusieurs sources
 * pour un même champ (ex : description = field_4 OU field_3 OU Title).
 * Renvoie '' si tous les candidats sont vides/null.
 *
 * `String(c).trim()` : gère le cas où c serait un nombre (0 → '0' → ''
 * après trim ne se produira pas). filter(Boolean) sur '0' renverrait true,
 * donc on utilise && c et trim() explicite.
 */
function pickFirstString(...candidates: Array<string | undefined | null>): string {
  for (const c of candidates) {
    if (c && String(c).trim()) return String(c).trim()
  }
  return ''
}

/**
 * Variante de pickFirstString qui passe ensuite par stripHtml.
 * Pour les champs dont le contenu peut être HTML (textes riches SharePoint).
 */
function pickFirstHtml(...candidates: Array<string | undefined | null>): string {
  const raw = pickFirstString(...candidates)
  return raw ? stripHtml(raw) : ''
}

/**
 * Découpe une chaîne en liste d'actions individuelles.
 *
 * Sépare sur tout ce qui ressemble à un séparateur de liste :
 *   - newline (\r\n ou \n)
 *   - point-virgule
 *   - puce (• – –)
 *
 * Filtre les chunks vides (ex : "; ; foo" → ['foo']).
 *
 * Use case : la cause/actions stockée dans field_4 peut être saisie en
 * format libre par les contrôleurs. On essaie de la transformer en liste
 * pour l'afficher en bullets dans le bulletin.
 */
function splitActions(raw: string): string[] {
  if (!raw) return []
  return raw
    .split(/\r?\n|;|•|–|–/)
    .map(s => s.trim())
    .filter(Boolean)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — CONSTRUCTION DE LA TIMELINE (CYCLE DE VIE)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Construit la timeline du ticket en agrégeant :
 *   1. Les 4 dates clés (déclaration, ouverture, régularisation, clôture)
 *   2. Les notes du journal field_23 si présent (logs horodatés)
 *
 * Logique :
 *   - Déclaration : field_0 (date métier) sinon Created (auto-SharePoint)
 *   - Ouverture   : dateOuvertureTicket sinon == déclaration (pas un step
 *     séparé si identique)
 *   - Régularisation : field_9 (date à laquelle le contrôleur a marqué
 *     l'anomalie comme corrigée)
 *   - Clôture     : date_cloture_ticket
 *
 * Journal field_23 :
 *   Format attendu : `[2026-05-01 10:30] Note libre`
 *   La regex extrait la date et le texte. Si pas de date détectée,
 *   l'entrée est ajoutée sans date (apparaît à la fin du tri).
 *
 * Tri final : chronologique croissant. Les entrées sans date (date = 0)
 * remontent en haut de la timeline.
 */
export function buildLifecycleSteps(ticket: DCPO_LISTE_ANORMALIERead): LifecycleStep[] {
  const ext = ticket as ExtendedTicket

  // ─── Étape 1 : Dates clés ──────────────────────────────────────────
  // Sémantique métier :
  //   - `Created` (horodatage SP auto à la soumission) → "Déclaration"
  //     (= moment où le déclarant a soumis le ticket dans le système).
  //   - `field_0` (date métier saisie dans le formulaire) → "Survenance"
  //     (= date à laquelle l'anomalie a réellement eu lieu sur le terrain).
  //
  // L'ordre d'affichage demandé n'est PAS chronologique : on veut
  // "Déclaration" puis "Survenance" en tête de timeline, puis les autres
  // étapes (ouverture du ticket, régularisation, clôture, journal) triées
  // chronologiquement à la suite.
  const systemCreated = parseDateSafe(ticket.Created)
  const survenance = parseDateSafe(ticket.field_0)
  const opened = parseDateSafe(ticket.dateOuvertureTicket)
  const regularized = parseDateSafe(ticket.field_9)
  const closed = parseDateSafe(ticket.date_cloture_ticket)

  // Étapes "en tête" : ordre fixe (Déclaration → Survenance), non triées.
  const head: LifecycleStep[] = []
  if (systemCreated) {
    head.push({
      type: 'declaration',
      label: 'Déclaration',
      date: systemCreated.toISOString(),
      description: ticket.declarant_anormalie?.DisplayName,
    })
  }
  if (survenance && survenance.getTime() !== systemCreated?.getTime()) {
    head.push({
      type: 'survenance',
      label: 'Survenance',
      date: survenance.toISOString(),
    })
  }

  // Étapes "queue" : triées chronologiquement entre elles.
  const tail: LifecycleStep[] = []
  // On évite un step "Ouverture" doublon si la date est == Survenance OU
  // == Déclaration (cas d'un ticket ouvert le jour de sa survenance).
  if (
    opened &&
    opened.getTime() !== survenance?.getTime() &&
    opened.getTime() !== systemCreated?.getTime()
  ) {
    tail.push({
      type: 'opening',
      label: 'Ouverture du ticket',
      date: opened.toISOString(),
    })
  }
  if (regularized) {
    tail.push({
      type: 'regularization',
      label: 'Régularisation',
      date: regularized.toISOString(),
      description: ticket.personneAffecter?.DisplayName,
    })
  }
  if (closed) {
    tail.push({
      type: 'closure',
      label: 'Clôture',
      date: closed.toISOString(),
    })
  }

  // ─── Étape 2 : Journal field_23 (si présent) ───────────────────────
  const journalRaw = ext.field_23
  if (journalRaw && typeof journalRaw === 'string') {
    const lines = stripHtml(journalRaw).split(/\r?\n+/).map(l => l.trim()).filter(Boolean)
    lines.forEach(line => {
      // Regex : capture les formats type "[YYYY-MM-DD HH:MM:SS] texte"
      // Groupe 1 : date ISO ; Groupe 2 : reste de la ligne
      const dateMatch = line.match(/^\[?(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?)\]?\s*[:\-–]?\s*(.*)$/)
      if (dateMatch) {
        tail.push({
          type: 'log',
          label: dateMatch[2] || 'Note',
          date: dateMatch[1],
        })
      } else {
        // Note sans date détectable
        tail.push({
          type: 'log',
          label: line,
        })
      }
    })
  }

  // ─── Étape 3 : Tri chronologique de la queue, head intouché ────────
  tail.sort((a, b) => {
    const da = a.date ? new Date(a.date).getTime() : 0
    const db = b.date ? new Date(b.date).getTime() : 0
    return da - db
  })

  return [...head, ...tail]
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 BIS — PARSING DES SECTIONS DE DESCRIPTION
 *
 * Pourquoi : les champs SP `field_4` (description) sont remplis par le
 * formulaire de clôture en CONCATENANT toutes les sections sur une seule
 * ligne, sous la forme :
 *   "Cause immédiate : <texte>Cause racine : <texte>Actions menées : <texte>Observations : <texte>"
 *
 * Les colonnes dédiées (causeImmediate, causeRacine, observations,
 * actionsAMener) restent souvent VIDES. Si on se contente du fallback en
 * cascade `pickFirstHtml(ext.causeImmediate, ext.field_11, …)`, on affiche
 * "—" alors que toute l'info est dans la description.
 *
 * Solution : on parse la description pour en extraire chaque section. Si une
 * section est trouvée, elle prend la priorité sur "—". Sinon on garde le
 * fallback existant.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Sections que l'on tente d'extraire de la description concaténée.
 * Synchronisé avec les libellés produits par le formulaire de clôture
 * (cf. Anomalies.tsx → buildResolutionDescription).
 */
const SECTION_LABELS = [
  'Cause immédiate',
  'Cause racine',
  'Actions menées',
  'Action menée',
  'Observations',
  'Observation',
] as const

/**
 * Parse une chaîne contenant les sections concaténées en clé/valeur.
 *
 * Algorithme :
 *   1. On construit une regex unique qui matche un libellé de section,
 *      suivi de ":" puis du texte jusqu'au PROCHAIN libellé (lookahead).
 *      Cela gère les sections collées sans séparateur ("…clientsCause racine").
 *   2. Normalise les clés pour matcher l'orthographe canonique (les pluriels
 *      "Action menée" / "Observations" sont uniformisés).
 *
 * Retourne un objet avec les sections trouvées (valeurs déjà trimées).
 */
function parseDescriptionSections(raw: string | null | undefined): Record<string, string> {
  if (!raw) return {}
  // Pattern : (libellé) \s* : \s* (texte jusqu'au prochain libellé ou fin)
  // Le `?:` rend le groupe alternative non-capturant pour ne pas polluer
  // les groupes nommés. Lookahead non-greedy → on s'arrête au prochain
  // libellé même collé sans espace.
  const labelsAlt = SECTION_LABELS.map(l => l.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')
  const re = new RegExp(
    `(${labelsAlt})\\s*:\\s*([\\s\\S]*?)(?=(?:${labelsAlt})\\s*:|$)`,
    'gi',
  )
  const out: Record<string, string> = {}
  let m: RegExpExecArray | null
  while ((m = re.exec(raw)) !== null) {
    const key = m[1].trim()
    const val = m[2].trim()
    if (!val) continue
    // Normalisation (singulier ↔ pluriel)
    let canonical = key
    if (/^action menée$/i.test(key)) canonical = 'Actions menées'
    else if (/^observation$/i.test(key)) canonical = 'Observations'
    // Si on a déjà rempli cette clé via un précédent match, on garde le
    // premier (cas d'une description qui répéterait un libellé).
    if (out[canonical] === undefined) out[canonical] = val
  }
  return out
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 7 — FONCTION PRINCIPALE : BULLETIN CONSOLIDÉ
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Construit un ConsolidatedBulletin complet à partir d'un ticket SharePoint
 * et des référentiels (agences, réseaux).
 *
 * C'est LA fonction principale du module. Appelée pour chaque ticket dans
 * la grille de bulletins.
 *
 * Stratégie :
 *   1. Parser toutes les dates en objets Date sécurisés
 *   2. Calculer le délai (déclaration → clôture, fallback régularisation)
 *   3. Résoudre les textes avec fallback en cascade
 *   4. Construire la timeline via buildLifecycleSteps
 *   5. Construire le numéro et le titre affichés en gros
 *   6. Retourner l'objet bulletin complet
 *
 * Note : retourne une nouvelle structure (pas une mutation du ticket).
 * → Les pages peuvent garder le ticket original si besoin.
 */
export function buildConsolidatedBulletin(
  ticket: DCPO_LISTE_ANORMALIERead,
  agences: DCPO_LISTE_AGENCESRead[],
  reseaux: DCPO_LISTE_RESEAUXRead[],
): ConsolidatedBulletin {
  const ext = ticket as ExtendedTicket

  // Dates parsées une seule fois pour réutilisation
  const declarationDate = parseDateSafe(ticket.field_0) ?? parseDateSafe(ticket.Created)
  const openingDate = parseDateSafe(ticket.dateOuvertureTicket) ?? declarationDate
  const regularizationDate = parseDateSafe(ticket.field_9)
  const closureDate = parseDateSafe(ticket.date_cloture_ticket)

  // Délai métier : préfère closureDate, fallback regularizationDate
  // (au cas où un ticket Resolu n'a pas encore date_cloture_ticket renseignée)
  const delayDays = diffInDays(declarationDate, closureDate ?? regularizationDate)

  // Fallback en cascade pour les textes principaux
  const descriptionFull = pickFirstHtml(ticket.field_4, ticket.field_3, ticket.Title)
  // Parsing des sections concaténées dans la description :
  // le formulaire de clôture remplit field_4 avec
  // "Cause immédiate : X Cause racine : Y Actions menées : Z Observations : W".
  // On extrait chaque section pour les afficher proprement dans le bulletin,
  // au cas où les colonnes dédiées (causeImmediate, etc.) seraient vides.
  const parsed = parseDescriptionSections(descriptionFull)
  // Si on a réussi à parser des sections, la "Description" affichée est
  // SEULEMENT le texte avant la première section (sinon on répète tout ce
  // qui sera déjà détaillé en dessous). Sinon, on garde tout.
  const firstLabelIdx = (() => {
    let min = -1
    for (const label of SECTION_LABELS) {
      const idx = descriptionFull.search(new RegExp(`\\b${label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*:`, 'i'))
      if (idx >= 0 && (min === -1 || idx < min)) min = idx
    }
    return min
  })()
  const description = firstLabelIdx >= 0
    ? descriptionFull.slice(0, firstLabelIdx).trim()
    : descriptionFull
  const causeImmediate =
    pickFirstHtml(ext.causeImmediate, ext.field_11) ||
    parsed['Cause immédiate'] ||
    ''
  const causeRacine =
    pickFirstHtml(ext.causeRacine, ext.field_12) ||
    parsed['Cause racine'] ||
    ''
  const observations =
    pickFirstHtml(ext.observationsBulletin, ext.field_22) ||
    parsed['Observations'] ||
    ''
  // actionsMenees : colonne SP dédiée (priorité absolue depuis la migration
  // récente — le formulaire de clôture y écrit directement les actions).
  // Fallback en cascade pour les anomalies historiques :
  //   1. ext.actionsAMener (ancienne colonne nommée)
  //   2. ext.field_21      (ancienne colonne numérotée)
  //   3. parsed['Actions menées'] : section extraite de la description
  //      concaténée — couvre les anomalies clôturées AVANT que la colonne
  //      `actionsMenees` n'existe.
  const actionsRaw =
    ticket.actionsMenees ||
    pickFirstHtml(ext.actionsAMener, ext.field_21) ||
    parsed['Actions menées'] ||
    ''
  const actions = splitActions(actionsRaw)
  const domaine = pickFirstString(ext.domaineAnomalie, ticket.domaineActivite, ext.field_13)
  const natureRisque = pickFirstString(ext.natureRisque, ext.field_14)
  // typeSanction : renseigné à la clôture par le contrôleur (cf. modale
  // de résolution dans Anomalies.tsx). Vide en cours de vie de l'anomalie.
  const typeSanction = ticket.typeSanction ?? ''
  const occurrences = ext.field_20 ?? undefined

  // {Link} = URL native SharePoint vers l'item dans l'UI moderne
  const sharePointUrl = ticket['{Link}'] ?? undefined

  // Numéro affiché : "T-{ID}" ou "—" si pas d'ID (ne devrait pas arriver)
  const numero = ticket.ID ? `T-${ticket.ID}` : '—'
  // Titre : Title officiel sinon field_3 sinon début de description
  const titre = pickFirstString(ticket.Title, ticket.field_3, description.substring(0, 80))

  return {
    ticket,
    numero,
    titre: titre || `Anomalie ${numero}`,
    agenceLabel: findAgenceLabel(agences, ticket.field_6),
    reseauLabel: findReseauLabel(reseaux, ticket.field_7),
    declarantName: ticket.declarant_anormalie?.DisplayName ?? '—',
    auteurName: ticket.auteur_anormalie?.DisplayName ?? '—',
    affecteName: ticket.personneAffecter?.DisplayName ?? '—',
    declarationDate,
    openingDate,
    regularizationDate,
    closureDate,
    delayDays,
    classification: ticket.field_5 ?? '',
    criticite: ticket.criticiteAnomalie ?? '',
    statut: ticket.field_10 ?? '',
    montant: ticket.field_8 ?? undefined,
    description,
    causeImmediate,
    causeRacine,
    observations,
    actions,
    domaine,
    natureRisque,
    typeSanction,
    occurrences,
    lifecycle: buildLifecycleSteps(ticket),
    sharePointUrl,
  }
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 8 — FILTRAGE & GUARDS
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Garde de filtrage : retourne true ssi le ticket est éligible au bulletin.
 * Utilisé après le fetch pour re-filtrer côté client (sécurité).
 *
 * Critère unique : statut field_10 ∈ ['Resolu', 'Clos'].
 */
export function isBulletinTicket(ticket: DCPO_LISTE_ANORMALIERead): boolean {
  if (!ticket.field_10) return false
  return BULLETIN_STATUS_VALUES.some(s => ticket.field_10 === s)
}

/**
 * Applique le jeu de filtres aux bulletins consolidés.
 *
 * Stratégie : un seul .filter() qui rejette dès qu'un critère ne matche pas
 * (early return false). Ordre des tests optimisé pour exclure rapidement
 * les non-matchs (champs simples avant search texte coûteux).
 *
 * Détails par filtre :
 *
 *   1. status        : compare b.statut directement (sauf 'Tous' = bypass)
 *   2. classification: comparaison stricte
 *   3. criticite     : comparaison stricte
 *   4. agence        : compare l'ID brut du ticket avec l'ID filtre
 *   5. agent (texte) : recherche partielle dans nom + email auteur
 *   6. affecte (texte): recherche partielle dans nom + email personneAffecter
 *   7. dates clôture : intervalle [closureFrom, closureTo] avec fallback
 *                      sur regularizationDate (anciens tickets sans clôture)
 *   8. search libre  : haystack = concat de tous les textes du bulletin,
 *                      puis includes() lowercase
 */
export function applyBulletinFilters(
  bulletins: ConsolidatedBulletin[],
  filters: BulletinFilters,
): ConsolidatedBulletin[] {
  // Pré-calcul : les opérations toLowerCase / parsing dates sont faites
  // UNE fois ici, pas N fois dans le .filter()
  const search = filters.search.trim().toLowerCase()
  const agentTerm = filters.agent.trim().toLowerCase()
  const affecteTerm = filters.affecte.trim().toLowerCase()
  const declarantTerm = filters.declarant.trim().toLowerCase()
  const fromDate = filters.dateFrom ? new Date(`${filters.dateFrom}T00:00:00`) : undefined
  const toDate = filters.dateTo ? new Date(`${filters.dateTo}T23:59:59`) : undefined

  return bulletins.filter(b => {
    // Filtres simples (égalité) — exclusion rapide
    if (filters.status !== 'Tous' && b.statut !== filters.status) return false
    if (filters.classification && b.classification !== filters.classification) return false
    if (filters.criticite && b.criticite !== filters.criticite) return false
    if (filters.agence && String(b.ticket.field_6 ?? '') !== filters.agence) return false
    if (filters.reseau && String(b.ticket.field_7 ?? '') !== filters.reseau) return false
    if (
      filters.domaineActivite &&
      (b.ticket.domaineActivite ?? '') !== filters.domaineActivite
    ) return false

    // Filtres texte (recherche partielle dans nom + email)
    if (agentTerm) {
      const haystack = `${b.auteurName} ${b.ticket.auteur_anormalie?.Email ?? ''}`.toLowerCase()
      if (!haystack.includes(agentTerm)) return false
    }
    if (affecteTerm) {
      const haystack = `${b.affecteName} ${b.ticket.personneAffecter?.Email ?? ''}`.toLowerCase()
      if (!haystack.includes(affecteTerm)) return false
    }
    if (declarantTerm) {
      const haystack = `${b.declarantName} ${b.ticket.declarant_anormalie?.Email ?? ''}`.toLowerCase()
      if (!haystack.includes(declarantTerm)) return false
    }

    // Filtre par intervalle de dates de clôture
    if (fromDate || toDate) {
      const ref = b.closureDate ?? b.regularizationDate
      if (!ref) return false  // pas de date dispo → exclu d'office
      if (fromDate && ref < fromDate) return false
      if (toDate && ref > toDate) return false
    }

    // Recherche libre — coûteuse, en dernier
    if (search) {
      const haystack = [
        b.numero,
        b.titre,
        b.description,
        b.declarantName,
        b.auteurName,
        b.affecteName,
        b.agenceLabel,
        b.reseauLabel,
        b.classification,
        b.criticite,
        b.causeImmediate,
        b.causeRacine,
        b.observations,
        ...b.actions,
      ]
        .filter(Boolean)
        .join(' ')
        .toLowerCase()
      if (!haystack.includes(search)) return false
    }
    return true
  })
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 9 — FORMATTERS POUR L'AFFICHAGE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Formate une Date en jj/mm/aaaa (locale FR).
 * Retourne "—" pour les dates absentes (UI plus propre que vide).
 */
export function formatDate(date?: Date | null): string {
  if (!date) return '—'
  return date.toLocaleDateString('fr-FR')
}

/**
 * Formate une Date OU une string ISO en date+heure FR.
 * Accepte les deux types pour faciliter l'usage côté JSX où on a parfois
 * une string brute (ex : `step.date` de la timeline).
 */
export function formatDateTime(date?: Date | string | null): string {
  if (!date) return ''
  const d = typeof date === 'string' ? new Date(date) : date
  if (Number.isNaN(d.getTime())) return ''
  return d.toLocaleString('fr-FR')
}

/**
 * Formate un montant numérique avec séparateurs de milliers (locale FR).
 * Ex : 1500000 → "1 500 000"
 *
 * Distinction null/undefined vs 0 : un montant à 0 est affiché "0", pas "—".
 */
export function formatAmount(value?: number | null): string {
  if (value === null || value === undefined) return '—'
  return value.toLocaleString('fr-FR')
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 10 — MAPPING VERS CLASSES CSS
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Mappe une criticité vers une classe CSS pour la pastille colorée.
 * Conversion case-insensitive (.toLowerCase()) pour tolérer les variantes
 * de saisie ("Critique", "CRITIQUE", "critique").
 *
 * Classes CSS définies dans Dashboard.css : .crit-faible / -moyenne / -haute
 * / -critique / -default
 */
export function getCriticiteClass(criticite?: string | null): string {
  switch ((criticite ?? '').toLowerCase()) {
    case 'critique': return 'crit-critique'
    case 'haute': return 'crit-haute'
    case 'moyenne': return 'crit-moyenne'
    case 'faible': return 'crit-faible'
    default: return 'crit-default'
  }
}

/**
 * Mappe un statut vers une classe CSS pour la pastille de statut.
 *
 * Cas particulier : on accepte 'resolu' ET 'résolu' (avec accent) pour
 * tolérer les éventuelles variations de saisie ou de migration de données.
 * Les deux convergent vers .status-resolu (pastille bleue Material Design 3).
 */
export function getStatusClass(status?: string | null): string {
  switch ((status ?? '').toLowerCase()) {
    case 'clos': return 'status-clos'
    case 'resolu':
    case 'résolu': return 'status-resolu'
    case 'en cours': return 'status-en-cours'
    case 'ouvert': return 'status-ouvert'
    default: return 'status-default'
  }
}
