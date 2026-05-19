/**
 * ============================================================================
 * SERVICE — PLAN DE CONTRÔLE (PLAN ANNUEL DCPO)
 * ============================================================================
 *
 * Couche d'accès aux données pour le module "Plan de Contrôle".
 *
 * Persistance ACTUELLE : localStorage (interim, le temps que la liste
 * SharePoint soit créée et que les services soient générés).
 *
 * Persistance CIBLE : liste SharePoint DCPO_PLAN_CONTROLE.
 *
 * Pour basculer vers SharePoint quand la liste sera en place :
 *   1. Lancer `npm run dev` (déclenche `pac code run` → régénère les services)
 *   2. Remplacer le corps des fonctions list/get/create par des appels
 *      au service généré (DCPO_PLAN_CONTROLEService).
 *      L'API publique (interfaces, signatures) NE doit PAS changer pour
 *      éviter d'impacter le composant React.
 *
 * Source des données : feuille "PLAN DE CONTROLE 2025" du fichier Excel
 * Tableau_de_bord_KPI_DCPO_2026-plan-de-controle.xlsx fourni par l'utilisateur.
 *
 * Modèle de planning mensuel :
 *   - 12 mois, chaque mois indexé par son code 2 chiffres ('01' à '12')
 *   - Stockage : array de codes mois (ex: ['01', '03', '06']) côté front
 *   - Sérialisation SP : codes joints par ';' (ex: "01;03;06") pour rentrer
 *     dans une simple colonne texte (évite de créer 12 colonnes Yes/No)
 * ============================================================================
 */


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 1 — TYPES DU DOMAINE
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Fréquence d'exécution d'un contrôle.
 *
 * Valeurs alignées sur la colonne FREQUENCE du fichier Excel source.
 * "Quotidiennene" (faute de frappe vue dans 1 cellule) sera normalisée en
 * 'Quotidienne' à l'import si on en a besoin un jour.
 */
export type ControleFrequence = 'Quotidienne' | 'Hebdomadaire' | 'Mensuelle' | 'Annuelle'

/**
 * Statut workflow d'un contrôle dans le plan.
 *
 *   À planifier : créé mais pas de mois d'exécution choisis
 *   Planifié    : mois choisis, en attente
 *   En cours    : exécution démarrée
 *   Réalisé     : terminé pour l'année
 *   Annulé      : retiré du plan (conservé pour historique)
 */
export type ControleStatus = 'À planifier' | 'Planifié' | 'En cours' | 'Réalisé' | 'Annulé'

/** Code mois : '01' à '12'. Format zero-padded pour tri lexicographique. */
export type MoisCode =
  | '01' | '02' | '03' | '04' | '05' | '06'
  | '07' | '08' | '09' | '10' | '11' | '12'

/**
 * Représentation d'un contrôle dans le plan annuel.
 *
 * Mapping avec le fichier Excel "PLAN DE CONTROLE 2025" :
 *   id                → Ordre (auto)
 *   libelle           → LIBELLE
 *   categorie         → ligne de section parent (header dans Excel)
 *   objectif          → Objectifs
 *   objectifChiffre   → Objectifs Chiffrés
 *   frequence         → FREQUENCE
 *   responsable       → RESPONSABLE
 *   moisPlanifies     → cases cochées dans les colonnes Jan-Déc (codes '01'..'12')
 *   annee             → millésime du plan
 *   statut            → workflow interne (pas dans Excel — ajouté pour suivi)
 */
export interface ControleEntry {
  id: string
  libelle: string
  categorie: string
  objectif: string
  objectifChiffre: string
  frequence: ControleFrequence
  responsable: string
  moisPlanifies: MoisCode[]
  annee: number
  statut: ControleStatus
  createdAt: string  // ISO timestamp pour tri/audit
}

/** Données minimales pour créer une entrée. */
export interface CreateControleInput {
  libelle: string
  categorie: string
  objectif: string
  objectifChiffre: string
  frequence: ControleFrequence
  responsable: string
  moisPlanifies: MoisCode[]
  annee: number
  statut: ControleStatus
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 2 — STOCKAGE LOCALSTORAGE (INTERIM)
 *
 * À remplacer par DCPO_PLAN_CONTROLEService dès que la liste SP existe.
 * ────────────────────────────────────────────────────────────────────────── */

const STORAGE_KEY = 'reportingDCPO.planControle.v1'

const readAll = (): ControleEntry[] => {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY)
    return raw ? (JSON.parse(raw) as ControleEntry[]) : []
  } catch {
    return []
  }
}

const writeAll = (entries: ControleEntry[]): void => {
  try {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(entries))
  } catch {
    /* ignore : quota / mode privé */
  }
}

function uid(): string {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 3 — RÉFÉRENTIELS (CATÉGORIES + RESPONSABLES + MOIS)
 *
 * Issus directement du fichier Excel. Codés en dur car ils ne changent
 * pas souvent ; pour les modifier, éditer ces constantes (ou, à terme,
 * créer une liste SP DCPO_PLAN_CONTROLE_CATEGORIES si le besoin émerge).
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Catégories (sections) du plan de contrôle.
 * Issues des lignes "header" du fichier Excel — un PAC peut être rattaché
 * à exactement une catégorie pour faciliter le regroupement à l'affichage.
 */
export const PLAN_CONTROLE_CATEGORIES: string[] = [
  "Contrôle préventif dans les unités d'exploitation",
  "Surveillance des opérations locales",
  "Renforcement du contrôle des activités du Comex",
  "Sécurisation des revenus de la banque",
  "Renforcement du suivi des opérations remarquables",
  "Renforcement de la surveillance des activités des directions centrales",
  "Renforcement de la surveillance des opérations (trésorerie)",
  "Évaluation contrôle des souscriptions aux produits",
  "Contrôle des opérations monétiques",
  "Contrôle des journées comptables",
  "Évaluation des contrôles des caisses",
  "Contrôle des existants",
  "Surveillance des engagements",
  "Réconciliation des comptes",
  "Révision des comptes",
  "Évaluation des activités externalisées",
  "Autre",
]

/**
 * Responsables observés dans le fichier source.
 * Liste indicative pour autocomplete/suggestions — le champ reste libre
 * (saisie d'un nom qui n'est pas dans la liste autorisée).
 */
export const PLAN_CONTROLE_RESPONSABLES_SUGGESTIONS: string[] = [
  'Tous les contrôleurs',
  'Contrôleurs réseau',
  'BOUE',
  'KAMEDA',
  'KENGUE',
  'KETATCHOU',
  'NJOYA',
  'NKADJIE/KENGUE',
  'NTCHUINTOUO / KAMENI',
  'SIMO / WELADJI',
  'TADOUNG',
  'TADOUNG / NJOYA',
  'TAKEUDEU',
  'TCHINDA',
  'WELADJI',
  'YONDJE',
]

/** Liste des 12 mois avec leur libellé court pour les headers de tableau. */
export const MOIS: { code: MoisCode; label: string }[] = [
  { code: '01', label: 'Jan' },
  { code: '02', label: 'Fév' },
  { code: '03', label: 'Mar' },
  { code: '04', label: 'Avr' },
  { code: '05', label: 'Mai' },
  { code: '06', label: 'Juin' },
  { code: '07', label: 'Juil' },
  { code: '08', label: 'Août' },
  { code: '09', label: 'Sept' },
  { code: '10', label: 'Oct' },
  { code: '11', label: 'Nov' },
  { code: '12', label: 'Déc' },
]

/** Options pour les selects de fréquence. */
export const FREQUENCE_OPTIONS: ControleFrequence[] = [
  'Quotidienne',
  'Hebdomadaire',
  'Mensuelle',
  'Annuelle',
]

/** Options pour les selects de statut. */
export const STATUS_OPTIONS: ControleStatus[] = [
  'À planifier',
  'Planifié',
  'En cours',
  'Réalisé',
  'Annulé',
]


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 4 — API PUBLIQUE
 *
 * Une fois passé en SharePoint :
 *   - listControles() → DCPO_PLAN_CONTROLEService.getAll({ filter: `annee eq ${year}` })
 *   - createControle() → service.create({ ...payload, moisPlanifies: codes.join(';') })
 *   - Lecture : parser moisPlanifies via raw.split(';').filter(...)
 * ────────────────────────────────────────────────────────────────────────── */

/** Liste tous les contrôles (tri stable : catégorie puis libellé). */
export function listControles(): ControleEntry[] {
  const all = readAll()
  return [...all].sort((a, b) => {
    if (a.annee !== b.annee) return b.annee - a.annee  // année récente en premier
    if (a.categorie !== b.categorie) return a.categorie.localeCompare(b.categorie)
    return a.libelle.localeCompare(b.libelle)
  })
}

/** Récupère un contrôle par ID. */
export function getControle(id: string): ControleEntry | undefined {
  return listControles().find(c => c.id === id)
}

/**
 * Crée une nouvelle entrée dans le plan de contrôle.
 *
 * Validation minimale (l'UI fait le détail) :
 *   - libelle obligatoire
 *   - categorie obligatoire
 *   - annee valide (entier > 0)
 *
 * Note : les mois planifiés peuvent être vides à la création
 * (statut = 'À planifier' implicite côté UI).
 */
export function createControle(input: CreateControleInput): ControleEntry {
  if (!input.libelle.trim()) throw new Error('Libellé obligatoire.')
  if (!input.categorie.trim()) throw new Error('Catégorie obligatoire.')
  if (!input.annee || input.annee <= 0) throw new Error("Année invalide.")
  const entry: ControleEntry = {
    id: uid(),
    ...input,
    libelle: input.libelle.trim(),
    categorie: input.categorie.trim(),
    objectif: input.objectif.trim(),
    objectifChiffre: input.objectifChiffre.trim(),
    responsable: input.responsable.trim(),
    createdAt: new Date().toISOString(),
  }
  const all = readAll()
  all.unshift(entry)
  writeAll(all)
  return entry
}


/* ──────────────────────────────────────────────────────────────────────────
 * SECTION 5 — HELPERS UI
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Mappe un statut vers une classe CSS (réutilise les pastilles manager-pill
 * du reste de l'app pour cohérence visuelle).
 */
export function getControleStatusClass(statut: ControleStatus): string {
  switch (statut) {
    case 'À planifier': return 'manager-pill-pending'
    case 'Planifié': return 'manager-pill'
    case 'En cours': return 'manager-pill-pending'
    case 'Réalisé': return 'manager-pill-realise'
    case 'Annulé': return 'manager-pill-reporte'
    default: return 'manager-pill'
  }
}
