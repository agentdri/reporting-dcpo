/**
 * ============================================================================
 * SERVICE — PLAN D'ACTION CORRECTIF (PAC)
 * ============================================================================
 *
 * Couche d'accès aux données pour le module "Plan d'Action Correctif".
 *
 * Persistance ACTUELLE : localStorage (interim, le temps que les listes
 * SharePoint soient créées et que les services soient générés).
 *
 * Persistance CIBLE : deux listes SharePoint
 *   1. DCPO_LISTE_DIRECTIONS (référentiel des directions concernées)
 *   2. DCPO_PAC                (plans d'action correctifs)
 *
 * Pour basculer vers SharePoint quand les listes seront en place :
 *   1. Lancer `npm run dev` (déclenche `pac code run` → régénère les services)
 *   2. Remplacer le corps des fonctions list/get/create par des appels
 *      aux services générés (DCPO_LISTE_DIRECTIONSService, DCPO_PACService).
 *      L'API publique (interfaces, signatures) NE doit PAS changer pour
 *      éviter d'impacter le composant React.
 *
 * Schéma des colonnes attendues côté SharePoint :
 *   - DCPO_LISTE_DIRECTIONS : Title (code), libelle, actif (Yes/No)
 *   - DCPO_PAC : Title (intitulé), sourcePac, dateCreation, descriptionProbleme,
 *                causeImmediate, causeRacine, actionsCorrectives,
 *                directionsConcernees (codes séparés par ';'),
 *                echeance, kpi, annee, responsable,
 *                statutPac (Choice: Exécutée / En cours / Non Exécutée),
 *                observations, derniereEvaluation, nouveauDelai, meoDcpo
 * ============================================================================
 */


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/** Statut workflow d'un PAC (correspond à la colonne `statutPac` côté SP). */
export type PacStatus = 'Exécutée' | 'En cours' | 'Non Exécutée'

/** Direction concernée (référentiel maintenu dans DCPO_LISTE_DIRECTIONS). */
export interface Direction {
  /** Code court (ex: 'DJC', 'DCE', 'DSI'). Sert d'identifiant métier. */
  code: string
  /** Libellé complet (ex: 'Direction Juridique et Conformité'). */
  libelle: string
  /** Si false, la direction est masquée des sélecteurs mais conservée. */
  actif: boolean
}

/**
 * Représentation d'un Plan d'Action Correctif côté front.
 *
 * Mapping avec la feuille Excel "PAC DCPO" :
 *   id                       → Ordre (auto)
 *   sourcePac                → Source PAC
 *   dateCreation             → Date
 *   intitule                 → Intitulés PAC
 *   descriptionProbleme      → Description du problème
 *   causeImmediate           → Causes immédiate
 *   causeRacine              → Causes racines (profondes/réelles)
 *   actionsCorrectives       → Détermination des actions correctives
 *   directionsConcernees     → Directions concernées (codes)
 *   echeance                 → Échéance
 *   kpi                      → KPI
 *   annee                    → ANNEE
 *   responsable              → Responsable de mise en œuvre
 *   statut                   → Statut anomalie
 *   observations             → Observations
 *   derniereEvaluation       → Dernière date d'Evaluation
 *   nouveauDelai             → NOUVEAU DELAI PROPOSE
 *   meoDcpo                  → MEO de la DCPO
 */
export interface Pac {
  id: string
  sourcePac: string
  dateCreation: string         // YYYY-MM-DD
  intitule: string
  descriptionProbleme: string
  causeImmediate: string
  causeRacine: string
  actionsCorrectives: string
  /** Liste de codes Direction (ex: ['DSI', 'DJC']). */
  directionsConcernees: string[]
  echeance: string             // YYYY-MM-DD
  kpi: string
  annee: number
  responsable: string
  statut: PacStatus
  observations?: string
  derniereEvaluation?: string  // YYYY-MM-DD
  nouveauDelai?: string        // YYYY-MM-DD
  meoDcpo?: string
  createdAt: string            // ISO timestamp (interne, pour tri/audit)
}

/** Données minimales pour créer un PAC. */
export interface CreatePacInput {
  sourcePac: string
  dateCreation: string
  intitule: string
  descriptionProbleme: string
  causeImmediate: string
  causeRacine: string
  actionsCorrectives: string
  directionsConcernees: string[]
  echeance: string
  kpi: string
  annee: number
  responsable: string
  statut: PacStatus
  observations?: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — STOCKAGE LOCALSTORAGE (INTERIM)
 *
 * À supprimer / remplacer dès que les listes SharePoint sont en place et
 * que les services générés sont disponibles.
 * ────────────────────────────────────────────────────────────────────────── */

const STORAGE_KEY_PAC = 'reportingDCPO.pac.v1'
const STORAGE_KEY_DIRECTIONS = 'reportingDCPO.directions.v1'

/**
 * Directions par défaut au premier chargement (vues dans le tableau de bord
 * Excel fourni). L'utilisateur peut ensuite les modifier/compléter via
 * upsertDirection — ou, idéalement, les gérer dans la liste SharePoint
 * DCPO_LISTE_DIRECTIONS une fois celle-ci créée.
 */
const DEFAULT_DIRECTIONS: Direction[] = [
  { code: 'DJC', libelle: 'Direction Juridique et Conformité', actif: true },
  { code: 'DCE', libelle: 'Direction Commercial Entreprises', actif: true },
  { code: 'DSI', libelle: 'Direction des Systèmes d\'Information', actif: true },
]

const readJSON = <T>(key: string, fallback: T): T => {
  try {
    const raw = window.localStorage.getItem(key)
    if (!raw) return fallback
    return JSON.parse(raw) as T
  } catch {
    return fallback
  }
}

const writeJSON = (key: string, value: unknown): void => {
  try {
    window.localStorage.setItem(key, JSON.stringify(value))
  } catch {
    /* ignore — quota dépassé / mode privé */
  }
}

/** Génère un identifiant unique court (interne). */
function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — API — DIRECTIONS
 *
 * Une fois passé en SharePoint :
 *   - listDirections() → DCPO_LISTE_DIRECTIONSService.getAll({ filter: "actif eq 1" })
 *   - upsertDirection() → create / update sur le même service
 *   - removeDirection() → soft delete : on met actif=false plutôt que de supprimer
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Liste les directions actives (triées par code).
 *
 * Au premier appel sur un nouvel utilisateur, on initialise le localStorage
 * avec DEFAULT_DIRECTIONS pour éviter une UI vide au démarrage. L'utilisateur
 * peut ensuite éditer librement.
 */
export function listDirections(activeOnly: boolean = true): Direction[] {
  const all = readJSON<Direction[]>(STORAGE_KEY_DIRECTIONS, [])
  if (all.length === 0) {
    writeJSON(STORAGE_KEY_DIRECTIONS, DEFAULT_DIRECTIONS)
    return activeOnly ? DEFAULT_DIRECTIONS.filter(d => d.actif) : DEFAULT_DIRECTIONS
  }
  const filtered = activeOnly ? all.filter(d => d.actif) : all
  return [...filtered].sort((a, b) => a.code.localeCompare(b.code))
}

/**
 * Crée OU met à jour une direction (clé = code, case-sensitive).
 *
 * Utilisé par l'écran de gestion des directions (à venir) — pour l'instant
 * il suffit de modifier DEFAULT_DIRECTIONS ci-dessus pour préconfigurer la
 * liste, ou d'appeler upsertDirection depuis la console dev.
 */
export function upsertDirection(direction: Direction): void {
  const all = readJSON<Direction[]>(STORAGE_KEY_DIRECTIONS, [])
  const idx = all.findIndex(d => d.code === direction.code)
  if (idx >= 0) all[idx] = direction
  else all.push(direction)
  writeJSON(STORAGE_KEY_DIRECTIONS, all)
}

/** Désactive une direction (soft delete — préserve les PAC qui la référencent). */
export function deactivateDirection(code: string): void {
  const all = readJSON<Direction[]>(STORAGE_KEY_DIRECTIONS, [])
  const idx = all.findIndex(d => d.code === code)
  if (idx >= 0) {
    all[idx] = { ...all[idx], actif: false }
    writeJSON(STORAGE_KEY_DIRECTIONS, all)
  }
}

/** Résout un libellé à partir d'un code de direction (fallback = code lui-même). */
export function getDirectionLabel(code: string): string {
  const all = readJSON<Direction[]>(STORAGE_KEY_DIRECTIONS, DEFAULT_DIRECTIONS)
  return all.find(d => d.code === code)?.libelle ?? code
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — API — PAC
 *
 * Une fois passé en SharePoint :
 *   - listPACs() → DCPO_PACService.getAll({ orderBy: ['Created desc'] })
 *                  + mapper item SP → Pac (parser directionsConcernees split ';')
 *   - createPAC() → DCPO_PACService.create({ ...payload, directionsConcernees: codes.join(';') })
 * ────────────────────────────────────────────────────────────────────────── */

/** Liste tous les PAC (du plus récent au plus ancien). */
export function listPACs(): Pac[] {
  const all = readJSON<Pac[]>(STORAGE_KEY_PAC, [])
  return [...all].sort((a, b) => (a.createdAt < b.createdAt ? 1 : -1))
}

/** Récupère un PAC par ID. */
export function getPAC(id: string): Pac | undefined {
  return listPACs().find(p => p.id === id)
}

/**
 * Crée un nouveau PAC.
 *
 * Validation minimale ici (l'UI fait la validation détaillée) :
 *   - intitule obligatoire
 *   - au moins une direction concernée
 *   - statut valide
 */
export function createPAC(input: CreatePacInput): Pac {
  if (!input.intitule.trim()) {
    throw new Error('Intitulé obligatoire.')
  }
  if (input.directionsConcernees.length === 0) {
    throw new Error('Au moins une direction concernée est requise.')
  }
  const pac: Pac = {
    id: uid(),
    ...input,
    createdAt: new Date().toISOString(),
  }
  const all = listPACs()
  all.unshift(pac)
  writeJSON(STORAGE_KEY_PAC, all)
  return pac
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — HELPERS UI
 * ────────────────────────────────────────────────────────────────────────── */

/** Liste fermée des statuts PAC (pour les selects de filtre / formulaire). */
export const PAC_STATUS_OPTIONS: PacStatus[] = ['En cours', 'Exécutée', 'Non Exécutée']

/**
 * Mappe un statut PAC vers une classe CSS (réutilise les pastilles "manager-pill"
 * déjà définies pour rester cohérent visuellement avec les autres modules).
 */
export function getPacStatusClass(statut: PacStatus): string {
  switch (statut) {
    case 'Exécutée': return 'manager-pill-realise'
    case 'En cours': return 'manager-pill-pending'
    case 'Non Exécutée': return 'manager-pill-reporte'
    default: return 'manager-pill'
  }
}
