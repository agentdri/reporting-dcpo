/**
 * ============================================================================
 * SERVICE DE GESTION DES REPORTINGS CONTROLEUR (Module 3)
 * ============================================================================
 *
 * Rôle :
 *   Encapsule toute la logique de persistance et de calcul liée aux rapports
 *   quotidiens des contrôleurs (saisie, brouillon, validation manager).
 *
 * Implémentation actuelle :
 *   Persistance via window.localStorage (mémoire navigateur). Les données ne
 *   transitent pas par le serveur ; chaque utilisateur voit ses propres
 *   rapports sur sa machine.
 *
 * Migration future prévue :
 *   La liste SharePoint « ACTIVITE_Controleurs » n'est pas encore générée
 *   comme datasource Power Apps (cf. .power/schemas). Quand elle le sera,
 *   les fonctions exportées ici (listReports, getReport, createReport,
 *   updateReport, validateReport, loadDraft, saveDraft, clearDraft)
 *   pourront être réécrites pour appeler le service SharePoint généré
 *   SANS modifier les pages qui les consomment (ControllerReporting.tsx,
 *   ControllerReportingList.tsx) — c'est l'intérêt d'avoir cette couche
 *   d'abstraction.
 * ============================================================================
 */


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE
 * Toutes les structures de données manipulées par le module.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Statut global d'un rapport quotidien :
 *   - 'Soumis'   : le contrôleur a envoyé son rapport, il attend validation
 *   - 'Réalisé'  : le manager a validé (le rapport est officiellement accepté)
 *   - 'Reporté'  : le manager a invalidé (le rapport a été rejeté avec motif)
 */
export type ActivityStatus = 'Soumis' | 'Réalisé' | 'Reporté'

/**
 * Une ligne unitaire dans un rapport — représente UNE action menée par le
 * contrôleur durant la journée.
 *
 * Champs :
 *   - id          : identifiant local unique (généré par uid()) ; sert de key React
 *   - domaine     : catégorie de l'activité (ex : 'Contrôle', 'Réunion'…)
 *   - objet       : sujet précis (ex : 'Anomalie #T-12', 'Visite agence Bessengue')
 *   - action      : description libre de ce qui a été fait (3-500 caractères)
 *   - duree       : durée en minutes (5-600)
 *   - observation : note optionnelle (commentaire libre)
 */
export interface ActivityLine {
  id: string
  domaine: string
  objet: string
  action: string
  duree: number
  observation?: string
}

/**
 * Trace horodatée d'une décision manager (validation OU invalidation).
 * Stockée dans validationHistory[] du rapport pour conserver l'historique
 * complet des allers-retours manager ↔ rapport.
 *
 * Champs :
 *   - date         : ISO timestamp du moment de la décision
 *   - manager      : nom affiché du manager (pour traçabilité humaine)
 *   - managerEmail : email du manager (pour traçabilité technique)
 *   - decision     : 'Réalisé' (validation) ou 'Reporté' (rejet)
 *   - motif        : raison du rejet (obligatoire si Reporté, libre si Réalisé)
 */
export interface ActivityValidationNote {
  date: string
  manager: string
  managerEmail?: string
  decision: 'Réalisé' | 'Reporté'
  motif?: string
}

/**
 * Rapport quotidien complet, persisté côté localStorage.
 * C'est l'objet "racine" — tout le reste y est rattaché.
 *
 * Champs principaux :
 *   - id                  : identifiant unique du rapport
 *   - controleurEmail     : qui a soumis (clé d'identification)
 *   - controleurName      : nom affiché du contrôleur
 *   - date                : date du jour rapporté (format YYYY-MM-DD)
 *
 * Lignes & calculs dérivés (recalculés via computeTotals()) :
 *   - lines               : tableau des actions de la journée
 *   - totalHeures         : somme des durées en heures (arrondi 2 décimales)
 *   - tempsOccupe         : pourcentage de la journée occupée (base 8h)
 *   - statutJournee       : Calme / Normal / Chargé / Très chargé
 *
 * Workflow :
 *   - statut              : Soumis → Réalisé/Reporté après décision manager
 *   - validationHistory   : pile de toutes les décisions managers passées
 *
 * Métadonnées optionnelles :
 *   - anomaliesDetectees  : compteur saisi par le contrôleur
 *   - lienRapport         : URL vers un document externe (rapport Excel, etc.)
 *   - observationsGlobales: texte libre commentaire global du contrôleur
 *   - attachments         : liste de pièces jointes (nom + url)
 *
 * Audit :
 *   - submittedAt         : date de la PREMIÈRE soumission
 *   - updatedAt           : date de la DERNIÈRE modification (re-écrit à chaque update)
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
  anomaliesDetectees?: number
  lienRapport?: string
  observationsGlobales?: string
  attachments?: { name: string; url: string }[]
  validationHistory: ActivityValidationNote[]
  submittedAt: string
  updatedAt: string
}

/**
 * Brouillon d'un rapport en cours de saisie — sauvegardé automatiquement
 * pour que le contrôleur ne perde rien s'il rafraîchit la page.
 *
 * Différences avec ActivityReport :
 *   - PAS de id final (peut être assigné si nécessaire pour réutiliser un id)
 *   - PAS de calculs dérivés (totalHeures, tempsOccupe…) : recalculés à chaque
 *     ouverture pour éviter de stocker des valeurs obsolètes
 *   - PAS de statut, validationHistory, etc. — c'est juste un brouillon
 *   - 1 seul brouillon actif par contrôleur (clé localStorage = email)
 */
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
 * SECTION 2 — CONSTANTES & HELPERS LOCALSTORAGE
 * Couche bas-niveau d'accès au stockage navigateur.
 * Toutes les fonctions encapsulent les try/catch (le localStorage peut être
 * désactivé en mode privé, ou plein, ou inaccessible — il faut donc être
 * défensif et ne JAMAIS laisser une erreur localStorage casser l'app).
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Clé localStorage où sont stockés TOUS les rapports (un seul tableau JSON).
 * Suffixe ".v1" : permet de migrer le format dans le futur sans casser les
 * données existantes (on créerait une clé .v2 avec migration).
 */
const STORAGE_KEY_REPORTS = 'reportingDCPO.activityReports.v1'

/**
 * Préfixe des clés brouillons. Le suffixe est l'email du contrôleur
 * (ex : "reportingDCPO.activityDraft.v1.jordy@xxx.com")
 * → 1 brouillon par contrôleur, isolé des autres.
 */
const STORAGE_KEY_DRAFT_PREFIX = 'reportingDCPO.activityDraft.v1.'

/**
 * Lecture brute. Retourne null si la clé n'existe pas OU si localStorage
 * lève une exception (mode privé Safari par exemple).
 */
const READ_RAW = (key: string): string | null => {
  try { return window.localStorage.getItem(key) } catch { return null }
}

/**
 * Écriture brute. Silencieuse en cas d'erreur — l'app continue de
 * fonctionner même si la persistance échoue (ex : quota dépassé).
 */
const WRITE_RAW = (key: string, value: string): void => {
  try { window.localStorage.setItem(key, value) } catch { /* ignore */ }
}

/**
 * Suppression brute. Silencieuse comme WRITE_RAW.
 */
const REMOVE_RAW = (key: string): void => {
  try { window.localStorage.removeItem(key) } catch { /* ignore */ }
}

/**
 * Génère un identifiant unique court et lisible.
 * Format : "{timestamp_base36}-{random_base36}"
 *   ex : "lvg5h3wq-a8f2k1"
 *
 * Pourquoi pas crypto.randomUUID() ?
 *   - randomUUID() est plus long (36 chars) et pas dispo sur tous les browsers
 *   - Date.now() en base 36 garantit l'unicité temporelle (≠ entre 2 appels)
 *   - 6 chars random suffisent pour gérer les collisions à la milliseconde
 *   - Lisible dans les outils dev (debug)
 */
function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — PERSISTANCE BAS-NIVEAU DES RAPPORTS
 * Lit / écrit le tableau complet des rapports dans localStorage.
 * Toujours résilient : retourne [] si le JSON est corrompu ou inexistant.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Désérialise tous les rapports depuis localStorage.
 *
 * Sécurités :
 *   1. Si la clé n'existe pas → []
 *   2. Si le JSON est invalide → catch → []
 *   3. Si la valeur n'est pas un tableau → []
 *
 * Le cast `as ActivityReport[]` est un acte de foi : on suppose que
 * le contenu écrit avant correspond bien à la forme attendue.
 * En cas de migration future, c'est ICI qu'il faudrait valider le shape.
 */
function readAllReports(): ActivityReport[] {
  const raw = READ_RAW(STORAGE_KEY_REPORTS)
  if (!raw) return []
  try {
    const parsed = JSON.parse(raw)
    if (Array.isArray(parsed)) return parsed as ActivityReport[]
    return []
  } catch {
    return []
  }
}

/**
 * Sérialise et écrase TOUT le tableau de rapports d'un coup.
 * Stratégie "read-modify-write" simple : pour ajouter/modifier un seul
 * rapport, on relit tout, on patche en mémoire, on réécrit tout.
 * Acceptable tant que le volume reste modéré (< 1000 rapports = < 500 KB).
 */
function writeAllReports(reports: ActivityReport[]): void {
  WRITE_RAW(STORAGE_KEY_REPORTS, JSON.stringify(reports))
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — API PUBLIQUE : LECTURE DES RAPPORTS
 * Fonctions consommées par les composants React.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Liste tous les rapports, triés du plus récent au plus ancien (par date).
 *
 * Le tri se fait sur `report.date` (string YYYY-MM-DD) — la comparaison
 * lexicographique fonctionne car le format ISO est ordonnable nativement.
 *
 * Retour : nouveau tableau (le sort() in-place ne pollue pas le storage car
 * on l'applique sur la copie retournée par readAllReports()).
 */
export function listReports(): ActivityReport[] {
  return readAllReports().sort((a, b) => (a.date < b.date ? 1 : -1))
}

/**
 * Récupère un rapport par son identifiant unique.
 * Retourne undefined si non trouvé (pas d'erreur, l'appelant gère).
 */
export function getReport(id: string): ActivityReport | undefined {
  return readAllReports().find(r => r.id === id)
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — CRÉATION D'UN RAPPORT
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Données minimales nécessaires pour créer un rapport.
 * Ce qui est OMIS volontairement par rapport à ActivityReport :
 *   - id, totalHeures, tempsOccupe, statutJournee : calculés / générés
 *   - statut : forcé à 'Soumis' à la création
 *   - validationHistory : démarre toujours à []
 *   - submittedAt, updatedAt : tamponnés à new Date().toISOString()
 *
 * → Permet à l'appelant de fournir UNIQUEMENT ce qui est métier.
 */
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

/**
 * Calcule les totaux dérivés à partir des lignes :
 *   - totalHeures   : somme des durées en heures (arrondie au centième)
 *   - tempsOccupe   : pourcentage d'une journée de 8h occupée (cap à 100%)
 *   - statutJournee : étiquette qualitative basée sur totalHeures
 *
 * Détails des calculs :
 *
 *   1. totalMinutes = somme des durées valides
 *      - Number.isFinite(l.duree) : exclut NaN, Infinity, -Infinity
 *      - Math.max(0, l.duree)     : ignore les valeurs négatives
 *
 *   2. totalHeures = totalMinutes / 60, arrondi à 2 décimales
 *      - * 100 puis / 100 = trick d'arrondi standard JavaScript
 *
 *   3. tempsOccupe = (totalMinutes / 480) * 100, capé à 100
 *      - 480 minutes = 8 heures = journée standard de référence
 *      - Math.min(100, …) : si > 8h, on plafonne à 100% (heures sup hors scope)
 *
 *   4. statutJournee : seuils en heures
 *      - < 4h    → Calme
 *      - 4 à 7h  → Normal
 *      - 7 à 9h  → Chargé
 *      - ≥ 9h    → Très chargé
 *      Ces seuils alimentent l'indicateur visuel (chip coloré) côté UI.
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

/**
 * Crée un nouveau rapport et le persiste en localStorage.
 *
 * Étapes :
 *   1. Calculer les totaux (computeTotals)
 *   2. Construire l'objet ActivityReport complet (avec id généré, statut
 *      par défaut 'Soumis', validationHistory vide, dates ISO)
 *   3. Lire l'état actuel de tous les rapports
 *   4. Y ajouter le nouveau (push, en queue)
 *   5. Tout réécrire dans localStorage
 *   6. Retourner le rapport créé (utile à l'appelant pour redirection,
 *      affichage de confirmation, etc.)
 *
 * Pourquoi push (et pas unshift) ?
 *   L'ordre d'insertion ne compte pas — listReports() trie par date à la
 *   lecture. Push = O(1) vs unshift = O(n).
 */
export function createReport(input: CreateReportInput): ActivityReport {
  const totals = computeTotals(input.lines)
  const now = new Date().toISOString()
  const report: ActivityReport = {
    id: uid(),
    controleurEmail: input.controleurEmail,
    controleurName: input.controleurName,
    date: input.date,
    lines: input.lines,
    totalHeures: totals.totalHeures,
    tempsOccupe: totals.tempsOccupe,
    statutJournee: totals.statutJournee,
    statut: 'Soumis',
    anomaliesDetectees: input.anomaliesDetectees,
    lienRapport: input.lienRapport,
    observationsGlobales: input.observationsGlobales,
    attachments: input.attachments,
    validationHistory: [],
    submittedAt: now,
    updatedAt: now,
  }
  const all = readAllReports()
  all.push(report)
  writeAllReports(all)
  return report
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 6 — MISE A JOUR D'UN RAPPORT
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Modifie un rapport existant en appliquant un patch partiel.
 *
 * Signature TypeScript :
 *   patch: Partial<Omit<ActivityReport, 'id'>>
 *   - Partial<…>  : tous les champs deviennent optionnels (on patche ce qu'on veut)
 *   - Omit<…, 'id'> : on interdit de modifier l'id (sécurité référentielle)
 *
 * Comportement :
 *   1. Cherche le rapport par id ; si introuvable → undefined
 *   2. Fusionne l'existant avec le patch (spread `...all[idx], ...patch`)
 *   3. Force updatedAt à maintenant (tamponnage automatique)
 *   4. Si patch.lines a été fourni, RECALCULE les totaux dérivés
 *      (sinon ils deviendraient incohérents avec les nouvelles lignes)
 *   5. Réécrit tout le tableau et retourne le rapport patché
 *
 * Note : retourne le NOUVEL objet (pas une mutation in-place de l'original).
 */
export function updateReport(id: string, patch: Partial<Omit<ActivityReport, 'id'>>): ActivityReport | undefined {
  const all = readAllReports()
  const idx = all.findIndex(r => r.id === id)
  if (idx === -1) return undefined
  const updated: ActivityReport = {
    ...all[idx],
    ...patch,
    updatedAt: new Date().toISOString(),
  }
  if (patch.lines) {
    const totals = computeTotals(patch.lines)
    updated.totalHeures = totals.totalHeures
    updated.tempsOccupe = totals.tempsOccupe
    updated.statutJournee = totals.statutJournee
  }
  all[idx] = updated
  writeAllReports(all)
  return updated
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 7 — VALIDATION MANAGER (ACTIONS WORKFLOW)
 * Le manager peut Valider (→ Réalisé) ou Invalider (→ Reporté) un rapport.
 * Chaque action est tracée dans validationHistory pour audit.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Données fournies par le manager au moment de sa décision.
 *
 * Note importante : `motif` est optionnel ICI au niveau type, mais l'UI
 * (ManagerDetailModal) impose qu'il soit non vide quand decision === 'Reporté'.
 * Ce contrat métier est appliqué côté React, pas ici, pour que le service
 * reste flexible.
 */
export interface ValidationInput {
  manager: string
  managerEmail?: string
  decision: 'Réalisé' | 'Reporté'
  motif?: string
}

/**
 * Applique une décision manager à un rapport.
 *
 * Effets :
 *   1. Construit une nouvelle ActivityValidationNote avec horodatage
 *   2. Met à jour le statut du rapport (Réalisé OU Reporté)
 *   3. APPEND la note à validationHistory (jamais d'écrasement → audit complet)
 *      → Si un manager invalide puis revalide plus tard, l'historique conserve
 *        les 2 traces avec dates et motifs.
 *   4. Tamponne updatedAt
 *
 * Pourquoi `target.validationHistory ?? []` ?
 *   Précaution : si un rapport ancien (avant l'introduction de ce champ)
 *   n'avait pas validationHistory, on initialise à [] avant d'append.
 */
export function validateReport(id: string, input: ValidationInput): ActivityReport | undefined {
  const all = readAllReports()
  const idx = all.findIndex(r => r.id === id)
  if (idx === -1) return undefined
  const note: ActivityValidationNote = {
    date: new Date().toISOString(),
    manager: input.manager,
    managerEmail: input.managerEmail,
    decision: input.decision,
    motif: input.motif,
  }
  const target = all[idx]
  const updated: ActivityReport = {
    ...target,
    statut: input.decision,
    validationHistory: [...(target.validationHistory ?? []), note],
    updatedAt: new Date().toISOString(),
  }
  all[idx] = updated
  writeAllReports(all)
  return updated
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 8 — GESTION DES BROUILLONS
 * Auto-sauvegarde de la saisie en cours pour ne rien perdre en cas de
 * rafraîchissement, fermeture d'onglet, ou changement de page.
 * 1 brouillon par contrôleur (clé localStorage différenciée par email).
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Charge le brouillon du contrôleur (s'il existe).
 *
 * Normalisation : .toLowerCase() sur l'email pour éviter les doublons
 * "user@x.com" vs "User@X.com" qui créeraient 2 entrées différentes.
 *
 * Retour : ActivityDraft trouvé, ou undefined si :
 *   - aucune entrée pour cet email
 *   - le JSON est corrompu (catch silencieux)
 */
export function loadDraft(controleurEmail: string): ActivityDraft | undefined {
  const raw = READ_RAW(STORAGE_KEY_DRAFT_PREFIX + controleurEmail.toLowerCase())
  if (!raw) return undefined
  try {
    return JSON.parse(raw) as ActivityDraft
  } catch {
    return undefined
  }
}

/**
 * Sauvegarde un brouillon (écrase l'éventuel précédent du même contrôleur).
 *
 * Tamponnage automatique de updatedAt — l'appelant n'a pas à le faire.
 * Note : `...draft` SUIVI de updatedAt → le champ updatedAt du draft passé
 * en argument est ignoré et remplacé.
 *
 * Appel typique : depuis ControllerReporting.tsx, dans un useEffect avec
 * debounce 600ms quand le formulaire change.
 */
export function saveDraft(draft: ActivityDraft): void {
  WRITE_RAW(STORAGE_KEY_DRAFT_PREFIX + draft.controleurEmail.toLowerCase(), JSON.stringify({
    ...draft,
    updatedAt: new Date().toISOString(),
  }))
}

/**
 * Supprime le brouillon du contrôleur.
 * Appelé après une soumission réussie OU sur "Réinitialiser le formulaire".
 */
export function clearDraft(controleurEmail: string): void {
  REMOVE_RAW(STORAGE_KEY_DRAFT_PREFIX + controleurEmail.toLowerCase())
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 9 — HELPERS UI
 * Petits utilitaires consommés directement par les composants pour éviter
 * la duplication de logique métier dans le JSX.
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Génère une nouvelle ligne vide (utilisée à l'initialisation du formulaire
 * ou quand le contrôleur clique "+ Ajouter une ligne").
 *
 * Valeurs par défaut :
 *   - id     : nouvel uid (key React stable même avant saisie)
 *   - duree  : 30 minutes (compromis raisonnable pour démarrer)
 *   - autres : chaînes vides
 */
export function makeEmptyLine(): ActivityLine {
  return { id: uid(), domaine: '', objet: '', action: '', duree: 30, observation: '' }
}

/**
 * Liste fermée des domaines proposés dans le select "Domaine" de la saisie.
 * Si on veut ajouter / retirer des catégories métier, c'est ICI uniquement.
 */
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

/**
 * Limites de validation pour les champs des lignes.
 * Utilisés à la fois par :
 *   - validateLines() (validation côté logique avant soumission)
 *   - les attributs HTML (maxLength, min, max) côté UI pour bloquer la saisie
 */
export const ACTIVITY_LINE_LIMITS = {
  actionMinLength: 3,    // une action décrite en moins de 3 caractères = inutile
  actionMaxLength: 500,  // au-delà, ce serait probablement plusieurs actions
  dureeMin: 5,           // 5 min = unité minimale métier
  dureeMax: 600,         // 600 min = 10h, plafond raisonnable d'une seule action
}

/**
 * Erreur de validation associée à un champ précis d'une ligne précise.
 * Permet à l'UI d'afficher l'erreur EXACTEMENT sous l'input fautif
 * (et d'appliquer le style aria-invalid="true" pour l'accessibilité).
 *
 * Champs :
 *   - index   : position de la ligne dans le tableau (-1 = erreur globale)
 *   - field   : quel champ de ActivityLine est en faute
 *   - message : texte d'erreur en français à afficher
 */
export interface LineValidationError {
  index: number
  field: keyof ActivityLine
  message: string
}

/**
 * Valide l'ensemble des lignes du rapport avant soumission.
 *
 * Règles appliquées :
 *
 *   0. Au moins UNE ligne doit exister (sinon erreur globale index = -1)
 *
 *   Pour chaque ligne :
 *     1. domaine non vide
 *     2. objet non vide (avec trim, donc espaces seuls = vide)
 *     3. action.length entre actionMinLength (3) et actionMaxLength (500)
 *        - !line.action gère le cas undefined/null
 *        - .trim() évite que "   " soit considéré comme valide
 *     4. duree entre dureeMin (5) et dureeMax (600), nombre fini
 *        - Number(line.duree) gère le cas où l'input HTML retourne string
 *        - !Number.isFinite(duree) attrape NaN, Infinity
 *
 * Retour : tableau d'erreurs (vide → tout est valide → on peut soumettre).
 *
 * Appelé depuis ControllerReporting.tsx au clic "Soumettre le rapport" ;
 * si retour non vide, l'UI affiche les messages SOUS chaque champ et
 * empêche l'envoi.
 */
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
