/**
 * ============================================================================
 * MODULE 1 — ANOMALIES (Liste, création, gestion ticket)
 * ============================================================================
 *
 * Page centrale de l'application. Concentre presque tout le cycle de vie
 * d'une anomalie : création, affectation, suivi, édition, clôture.
 *
 * VUE PRINCIPALE
 * --------------
 *   - Tableau paginé des anomalies EN COURS (statuts Ouvert / En cours)
 *     UNIQUEMENT. Les Clos / Résolu ne s'affichent JAMAIS ici — elles sont
 *     consultables sur "Fiche récapitulatif de l'anomalie" (AnomalyBulletins).
 *     Cf. buildFilter() qui ajoute `field_10 ne 'Clos' and ne 'Resolu'` par
 *     défaut.
 *   - Colonnes toujours visibles : Numéro (T-{ID SP}), Statut, Date, Criticité,
 *     Délai (pastille colorée selon dateAffection + delai), [Actions].
 *   - Colonnes dépliables ("+/-") : Déclarant, Auteur, Personne affectée,
 *     Cause, Classification, Agence, Réseau, Montant, Date régularisation.
 *   - Cellules sticky (Numéro à gauche, Actions à droite) pour le scroll
 *     horizontal.
 *   - Filtres serveur (OData) : dates, agence, réseau, classification,
 *     criticité, domaine d'activité, statut.
 *   - Filtres client : agent (auteur), personne affectée (recherche libre).
 *   - Numéro de ticket = ID SharePoint (T-{item.ID}) → cohérent avec la
 *     page Fiche récapitulatif (un même ticket porte le MÊME numéro partout).
 *
 * MODALES (état piloté par variables dédiées, ouverture mutuellement exclusive)
 * --------------------------------------------------------------------------
 *   1. DÉTAIL (detailItem)
 *      - Lecture par défaut, mode édition au clic "Modifier" (managers ET
 *        contrôleur affecté à l'anomalie — cf. canEditAnomaly)
 *      - Édition : tous les champs métier sauf les Personne (auteur,
 *        déclarant, personne affectée) qui passent par leurs flux dédiés
 *
 *   2. AFFECTATION (affectItemId) — RÉSERVÉE AUX MANAGERS (canAffect)
 *      Workflow en 2 étapes :
 *        a) Rechercher un utilisateur Office 365 et le sélectionner
 *        b) Renseigner délai (défaut 3 jours) + commentaire d'affectation
 *      Écrit en SP : personneAffecter, delai, commentaireAffectation,
 *      dateAffection (horodatage) — cf. confirmAffect.
 *
 *   3. CHANGEMENT DE STATUT RAPIDE (statusItem)
 *      Bascule Ouvert ↔ En cours sans passer par la modale de résolution
 *      complète. Si statut = Resolu/Clos, ouvre aussi la date associée.
 *
 *   4. SUIVI DU TICKET (ticketItem)
 *      Vue synthétique des infos clés + commentaire d'affectation visible
 *      si renseigné. Sert d'écran de transit (Détail / Statut / Clôture).
 *
 *   5. CLÔTURE DE LA RÉSOLUTION (resolutionItem)
 *      Formulaire complet pour passer une anomalie en Resolu/Clos :
 *      statut final, dates, auteur, causes immédiate/racine, actions menées,
 *      observations, pièce jointe, mode de traitement (typeSanction).
 *      → Crée aussi une description structurée concaténée dans field_4.
 *
 * RÈGLES MÉTIER CLÉS
 * ------------------
 *   - VISIBILITÉ (côté serveur, cf. buildFilter) :
 *       Controleur → ne voit QUE les anomalies dont il est affecté
 *         (personneAffecter/Email eq userEmail)
 *       Manager (Chef_Departement, Directeur) → voit tout
 *
 *   - AFFECTATION :
 *       Bouton "Affecter" et champ "Personne affectée" du formulaire de
 *       création MASQUÉS pour les contrôleurs (cf. canAffect).
 *
 *   - ÉDITION INLINE (modale Détail) :
 *       Controleur peut éditer UNIQUEMENT les anomalies qui lui sont
 *       affectées (canEditAnomaly). Managers éditent tout.
 *
 *   - DÉLAI / ÉCHÉANCE :
 *       echeance = dateAffection + delai jours
 *       Colonne "Délai" affiche une pastille colorée selon diff
 *       (cf. getDelaiStatus) : vert si en cours, orange si échéance jour J,
 *       rouge si dépassée.
 *
 *   - CYCLE DE VIE :
 *       Création → Ouvert → (Affectation, En cours…) → Resolu / Clos
 *       Statut Clos = ticket fermé définitivement. La clôture ouvre la
 *       modale résolution qui demande tous les détails métier.
 *
 * PERSISTANCE
 * -----------
 *   - DCPO_LISTE_ANORMALIE via le service généré
 *   - Pièces jointes : workflow Power Automate qui retourne une URL,
 *     stockée dans urlPieceJointe (multi-URLs concaténées par " | ")
 *   - Mapping des field_N → cf. interface FormState ci-dessous
 * ============================================================================
 */

import { useState, useEffect, useMemo, useRef } from 'react'
import { useStagger } from '../lib/useGsap'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { User } from '../generated/models/Office365UsersModel'
import { appendUrl, getTicketAttachments, getAttachmentIcon } from '../lib/ticketAttachments'
import { formatMontantCompact, formatDateOnlyFR } from '../lib/formatters'
import { DOMAINE_ACTIVITE_OPTIONS, TYPE_SANCTION_OPTIONS } from '../lib/referentiels'
import { notifyAffectation, notifyAuteurAnomalie } from '../lib/teamsNotifications'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { ExportButtons } from '../components/ExportButtons'
import { findAgenceLabel, findReseauLabel, loadAgences, loadReseaux } from '../lib/spReferenceRows'
import { formatDateForExport } from '../lib/exporters'
import { getAllPages } from '../lib/sharePointPaging'
import { ModalOverlay } from '../components/ModalOverlay'

/**
 * Props passées par Dashboard.tsx — identité de l'utilisateur courant.
 * Utilisée pour pré-remplir le déclarant dans le formulaire de création.
 */
interface AnomaliesProps {
  userName?: string
  userEmail?: string
  /**
   * Rôle métier de l'utilisateur (issu de DCPO_LISTE_USER.fonction).
   *
   * Pilote les permissions d'affectation :
   *   - 'Controleur'                          → NE PEUT PAS affecter une anomalie
   *   - 'Chef_Departement' / 'Directeur'      → PEUVENT affecter
   *
   * Concrètement, quand le rôle vaut 'Controleur' :
   *   - le bouton "Affecter" est masqué dans le tableau d'anomalies
   *   - le champ "Personne affectée" est masqué dans le formulaire de création
   *
   * Note : cette restriction est UX/côté client uniquement. SharePoint peut
   * toujours rejeter une affectation si l'utilisateur n'a pas les permissions
   * effectives sur la colonne (couche de défense supplémentaire).
   */
  userRole?: string
}

/**
 * Forme du state du formulaire de création.
 *
 * Mapping des champs SharePoint :
 *   - field_0 : Date de l'anomalie (YYYY-MM-DD)
 *   - field_4 : Cause / description (texte HTML)
 *   - field_5 : Classification (Operationnel / Fraude / Commercial)
 *   - field_6 : Agence ID (string)
 *   - field_7 : Réseau ID (déduit de l'agence)
 *   - field_8 : Montant (nombre)
 *   - field_9 : Date de régularisation
 *   - field_10 : Statut (Ouvert / En cours / Resolu / Clos)
 *   - criticiteAnomalie : Faible / Moyenne / Haute / Critique
 */
interface FormState {
  field_0: string
  field_4: string
  field_5: string
  field_6: string
  field_7: string
  field_8: number
  field_9: string
  field_10: string
  criticiteAnomalie: string
  /** Domaine d'activité concerné (colonne SP `domaineActivite`). */
  domaineActivite: string
}

/** État initial vide du formulaire (utilisé après création / au reset). */
const EMPTY_FORM: FormState = {
  field_0: '',
  field_4: '',
  field_5: '',
  field_6: '',
  field_7: '',
  field_8: 0,
  field_9: '',
  field_10: '',
  criticiteAnomalie: '',
  domaineActivite: '',
}

/** Niveaux de criticité d'anomalie. Utilisé dans les selects. */
const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']

/**
 * Classification métier de l'anomalie (champ field_5).
 *
 * Référentiel officiel DCPO 2026 — synchronisé avec la colonne Choix
 * "Classification" de la liste SharePoint DCPO_LISTE_ANORMALIE.
 *
 * Pour ajouter une nouvelle valeur : la créer côté SP (Choice → Add value)
 * ET l'ajouter ici. Tant que les deux ne sont pas synchros, la sélection ne
 * sera pas écrite correctement (SP rejette les valeurs inconnues).
 */
const CLASSIFICATION_OPTIONS = [
  'Fraude interne',
  'Exécution, livraison et gestion des processus',
  'Fraude externe',
  'Interruptions de l\'activité et dysfonctionnements des systèmes',
  'Pratiques en matière d\'emploi et de sécurité du travail',
  'Clients, produits et pratiques commerciales',
  'Dommages occasionnés aux actifs physiques',
] as const

// DOMAINE_ACTIVITE_OPTIONS et TYPE_SANCTION_OPTIONS sont désormais centralisés
// dans lib/referentiels.ts (partagés avec d'autres modules).

/* ──────────────────────────────────────────────────────────────────────────
 * URL DU WORKFLOW POWER AUTOMATE (UPLOAD PIÈCE JOINTE)
 * Identique à ticketAttachments.ts — duplication historique pour le
 * uploadAttachment() local utilisé à la création. Pourrait être dédupliqué.
 * ────────────────────────────────────────────────────────────────────────── */
//const ATTACHMENT_API_URL = 'https://e78a17afcaf0e888989bbeca000173.f8.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/19d148d0144041f49ec16f59e318d7db/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=0Yeg1G81xgXL1XUkBAyoHpqdtpFq1PbZrJJPLmixw1M'
const ATTACHMENT_API_URL = 'https://default2bd82a682c7d4c43b0809b064410f6.cf.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/a8ced63bd1314a7897003b97eebd85e9/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=EdDW19nkBl6pWigNfp0OQPaiSIjBwyhAuWeTO3s8b8E'

/** Réponse du workflow Power Automate après upload. */
interface UploadResponse {
  success: boolean
  sharePointId?: number
  attachmentUrl?: string
  message?: string
}


/* ──────────────────────────────────────────────────────────────────────────
 * HELPERS DE FORMATAGE (UTILISÉS DANS LE COMPOSANT)
 * ────────────────────────────────────────────────────────────────────────── */

/**
 * Nettoie un texte HTML pour ne garder que le texte, en PRÉSERVANT les
 * sauts de ligne sémantiques (paragraphes, <br>, items de liste).
 *
 * Pourquoi : `textContent` natif colle les blocs sans séparation, ce qui
 * produit "Cause immédiate : RAS Cause racine : ..." au lieu d'une mise
 * en forme lisible.
 *
 * Étapes :
 *   1. Insérer un '\n' avant les balises de fermeture/sauts blocks-level
 *      pour matérialiser le saut visuel.
 *   2. Extraire le textContent → on a maintenant des vrais retours à la ligne.
 *   3. Heuristique métier : si le texte contient les libellés de section
 *      d'une description anomalie (Cause immédiate, Cause racine, Actions
 *      menées, Observations) collés sans séparation, on les remet sur des
 *      lignes distinctes.
 *   4. Compactage : remplace les enchaînements de 3+ newlines par 2 pour
 *      éviter de trop espacer.
 */
function stripHtml(html: string): string {
  if (!html) return ''
  // Étape 1 : matérialiser les blocs en sauts de ligne avant le parse
  const withBreaks = html
    .replace(/<\s*br\s*\/?\s*>/gi, '\n')
    .replace(/<\/\s*(p|div|li|h[1-6]|tr)\s*>/gi, '\n')
  const doc = new DOMParser().parseFromString(withBreaks, 'text/html')
  let text = doc.body.textContent ?? ''

  // Étape 3 : sections métier des descriptions d'anomalie. On insère un
  // saut de ligne AVANT chaque libellé (sauf en début de chaîne) si elles
  // sont collées à la phrase précédente. La regex tolère espaces et
  // ponctuation autour du ":".
  const SECTION_LABELS = [
    'Cause immédiate',
    'Cause racine',
    'Actions menées',
    'Action menée',
    'Observations',
    'Observation',
  ]
  for (const label of SECTION_LABELS) {
    // (?<!^|\n)  : pas en début de texte ou de ligne (lookbehind)
    // \s* avant : tolère espaces existants
    // \s*:      : tolère espaces avant ":"
    const re = new RegExp(`(?<!^|\\n)\\s*${label}\\s*:`, 'g')
    text = text.replace(re, `\n${label} :`)
  }

  // Étape 4 : nettoyage des sauts excessifs
  text = text
    .replace(/[ \t]+/g, ' ')     // espaces multiples → un seul
    .replace(/\n{3,}/g, '\n\n')  // 3+ newlines → 2
    .trim()
  return text
}

/**
 * État d'avancement d'un délai d'affectation d'anomalie.
 *
 * Calculé à partir de la date d'affectation, du délai (en jours) et du
 * jour courant. Utilisé pour la colonne "Délai" du tableau principal.
 *
 *   echeance = dateAffection + delai
 *   diff     = echeance - aujourd'hui (en jours arrondis à l'inférieur)
 *
 * Trois cas :
 *   - kind 'na'       : pas d'affectation ou délai invalide → "—"
 *   - kind 'restant'  : diff > 0 → "X j restant(s)" (vert)
 *   - kind 'echu'     : diff === 0 → "Échéance aujourd'hui" (orange)
 *   - kind 'depasse'  : diff < 0 → "Dépassée de X j" (rouge)
 */
interface DelaiStatus {
  kind: 'na' | 'restant' | 'echu' | 'depasse'
  diff: number
  label: string
}

function getDelaiStatus(
  dateAffection: string | undefined,
  delai: string | undefined,
): DelaiStatus {
  if (!dateAffection || !delai) return { kind: 'na', diff: 0, label: '—' }
  const delaiNum = parseInt(delai, 10)
  if (!Number.isFinite(delaiNum) || delaiNum <= 0) {
    return { kind: 'na', diff: 0, label: '—' }
  }
  const start = new Date(dateAffection)
  if (Number.isNaN(start.getTime())) return { kind: 'na', diff: 0, label: '—' }

  // Normalisation à minuit local pour comparer en jours pleins, indépendamment
  // de l'heure d'affectation et de l'heure courante du navigateur.
  const startDay = new Date(start.getFullYear(), start.getMonth(), start.getDate())
  const echeance = new Date(startDay)
  echeance.setDate(echeance.getDate() + delaiNum)
  const today = new Date()
  const todayDay = new Date(today.getFullYear(), today.getMonth(), today.getDate())

  const ONE_DAY = 86_400_000
  const diff = Math.round((echeance.getTime() - todayDay.getTime()) / ONE_DAY)

  if (diff > 0) return { kind: 'restant', diff, label: `${diff} j restant${diff > 1 ? 's' : ''}` }
  if (diff === 0) return { kind: 'echu', diff, label: "Échéance aujourd'hui" }
  const abs = Math.abs(diff)
  return { kind: 'depasse', diff, label: `Dépassée de ${abs} j${abs > 1 ? 's' : ''}` }
}

/**
 * Convertit un email en format "Claims" SharePoint pour les champs de type
 * Personne/Groupe.
 *
 * Format SharePoint : `i:0#.f|membership|email@domaine.com`
 *   - i:0    : claim provider type
 *   - #      : separator
 *   - .f     : federated identity
 *   - membership : provider name
 *
 * Indispensable pour persister un User dans une colonne SP People.
 */
function toClaims(email: string) {
  return `i:0#.f|membership|${email}`
}

/**
 * Convertit un File en string base64 (pour POST JSON vers Power Automate).
 * Cf. ticketAttachments.ts pour explication détaillée.
 */
function fileToBase64(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => {
      const result = reader.result as string
      const base64 = result.includes(',') ? result.split(',')[1] : result
      resolve(base64)
    }
    reader.onerror = () => reject(reader.error)
    reader.readAsDataURL(file)
  })
}

/**
 * Échappe les caractères HTML spéciaux pour empêcher l'injection XSS
 * quand on stocke du texte utilisateur dans un champ HTML SharePoint
 * (ex : field_4 description riche).
 *
 * Use case : buildResolutionHtml utilise cette fonction sur les inputs
 * du contrôleur avant de les envelopper en <p>...</p> pour SharePoint.
 */
function escapeHtml(input: string): string {
  return input
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

/**
 * Mappe un statut d'anomalie vers la classe CSS du badge ticket
 * (couleurs Material Design 3, cf. Dashboard.css).
 */
function statusBadgeClass(status?: string | null): string {
  switch ((status ?? '').toLowerCase()) {
    case 'ouvert': return 'ticket-badge-ouvert'
    case 'en cours': return 'ticket-badge-en-cours'
    case 'resolu': return 'ticket-badge-resolu'
    case 'clos': return 'ticket-badge-clos'
    default: return 'ticket-badge-default'
  }
}

/**
 * Mappe un statut vers une icône Unicode représentative pour le badge ticket.
 *   Ouvert  → ⚠ (warning)
 *   En cours → ⏳ (sablier)
 *   Resolu  → ✓ (check simple)
 *   Clos    → ✔ (check renforcé)
 */
function statusIcon(status?: string | null): string {
  switch ((status ?? '').toLowerCase()) {
    case 'ouvert': return '⚠'
    case 'en cours': return '⏳'
    case 'resolu': return '✓'
    case 'clos': return '✔'
    default: return '•'
  }
}

export default function Anomalies({ userName, userEmail, userRole }: AnomaliesProps) {
  /**
   * Permission "affecter une anomalie".
   *
   * Règle métier :
   *   - Chef_Departement / Directeur → peuvent affecter N'IMPORTE quelle anomalie
   *   - Controleur                   → peut affecter UNIQUEMENT les anomalies
   *                                    qu'il a DÉCLARÉES (declarant_anormalie.Email
   *                                    match userEmail, comparaison
   *                                    case-insensitive). Cas d'usage :
   *                                    un contrôleur qui a saisi une déclaration
   *                                    peut la réaffecter à un collègue plus
   *                                    compétent sans passer par un manager.
   *   - Rôle inconnu / non renseigné → permissif (true) pour ne pas bloquer
   *                                    en cas de configuration manquante
   */
  const canAffectAnomaly = (item: DCPO_LISTE_ANORMALIERead | null): boolean => {
    if (!item) return false
    if (userRole === 'Chef_Departement' || userRole === 'Directeur') return true
    if (userRole === 'Controleur') {
      const myEmail = userEmail?.toLowerCase()
      const declarantEmail = item.declarant_anormalie?.Email?.toLowerCase()
      return !!myEmail && declarantEmail === myEmail
    }
    return true
  }

  /**
   * Vérifie si l'utilisateur peut éditer une anomalie donnée dans la modale Détail.
   *
   * Règle métier :
   *   - Chef_Departement / Directeur → peuvent éditer N'IMPORTE quelle anomalie
   *   - Controleur                   → peut éditer si :
   *                                      • il est la personne affectée, OU
   *                                      • il est le déclarant de l'anomalie
   *                                    (la propriété est attachée à
   *                                    l'utilisateur qui a soumis le
   *                                    formulaire de création — cf.
   *                                    declarant_anormalie). Cas d'usage :
   *                                    le contrôleur veut corriger une faute
   *                                    de saisie sur sa propre anomalie.
   *   - Autre rôle / pas de rôle     → permissif (true) pour ne pas bloquer
   *                                    en cas de configuration manquante
   *
   * Le calcul est case-insensitive sur l'email pour absorber les variations
   * de casse SharePoint vs Office 365.
   *
   * @param item Anomalie à tester
   * @returns true si l'utilisateur peut basculer en mode édition
   */
  const canEditAnomaly = (item: DCPO_LISTE_ANORMALIERead | null): boolean => {
    if (!item) return false
    if (userRole === 'Chef_Departement' || userRole === 'Directeur') return true
    if (userRole === 'Controleur') {
      const myEmail = userEmail?.toLowerCase()
      if (!myEmail) return false
      const affecteEmail = item.personneAffecter?.Email?.toLowerCase()
      const declarantEmail = item.declarant_anormalie?.Email?.toLowerCase()
      return affecteEmail === myEmail || declarantEmail === myEmail
    }
    // Rôle inconnu : permissif (cf. note canAffect).
    return true
  }

  /**
   * Vérifie si l'utilisateur peut CLORE une anomalie (passer à Resolu/Clos
   * ou ouvrir la modale de résolution complète).
   *
   * Règle métier :
   *   - Chef_Departement / Directeur → peuvent clore N'IMPORTE quelle anomalie
   *   - Controleur                   → peut clore UNIQUEMENT les anomalies
   *                                    qui lui sont AFFECTÉES (un contrôleur
   *                                    n'est pas autorisé à clore le travail
   *                                    d'un collègue)
   *   - Rôle inconnu                 → permissif (cohérent avec canAffect /
   *                                    canEditAnomaly)
   *
   * Logiquement équivalent à canEditAnomaly aujourd'hui, mais conservé en
   * helper séparé pour pouvoir diverger plus tard (par ex : un contrôleur
   * pourrait éditer mais pas clore, ou inversement, selon les évolutions
   * métier).
   */
  const canCloseAnomaly = (item: DCPO_LISTE_ANORMALIERead | null): boolean => {
    if (!item) return false
    if (userRole === 'Chef_Departement' || userRole === 'Directeur') return true
    if (userRole === 'Controleur') {
      const affecteEmail = item.personneAffecter?.Email?.toLowerCase()
      const myEmail = userEmail?.toLowerCase()
      return !!affecteEmail && !!myEmail && affecteEmail === myEmail
    }
    return true
  }

  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS — DONNÉES PRINCIPALES & FORMULAIRE DE CRÉATION
   * ────────────────────────────────────────────────────────────────────── */

  /** Liste brute des anomalies récupérées du serveur (filtres OData appliqués). */
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  /** Indicateur de chargement (réseau en cours). */
  const [loading, setLoading] = useState(true)
  /** Visibilité du formulaire de création (toggle au clic "+ Nouvelle anomalie"). */
  const [showForm, setShowForm] = useState(false)
  /** State du formulaire de création (cf. FormState). */
  const [form, setForm] = useState(EMPTY_FORM)
  /** True pendant l'envoi de la création (désactive le bouton Submit). */
  const [submitting, setSubmitting] = useState(false)
  /** Pièce jointe sélectionnée à la création (uploadée après création). */
  const [attachment, setAttachment] = useState<File | null>(null)


  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS — FILTRES
   *
   * Deux groupes :
   *   1. "filter*"  : valeurs liées aux inputs (changent à la frappe)
   *   2. "applied*" : snapshot pris au clic sur "Rechercher" — c'est ce
   *      qui pilote le filtrage effectif (côté client) ou le rechargement
   *      du serveur. Permet de découpler l'UX (saisie libre) de l'exécution
   *      (uniquement au clic).
   * ────────────────────────────────────────────────────────────────────── */

  // Filtres serveur (appliqués via OData $filter au clic "Rechercher")
  const [filterDateFrom, setFilterDateFrom] = useState('')
  const [filterDateTo, setFilterDateTo] = useState('')
  const [filterAgence, setFilterAgence] = useState('')
  const [filterReseau, setFilterReseau] = useState('')
  const [filterClassification, setFilterClassification] = useState('')
  const [filterCriticite, setFilterCriticite] = useState('')
  /** Filtre par domaine d'activité (colonne SP `domaineActivite`). */
  const [filterDomaineActivite, setFilterDomaineActivite] = useState('')
  /**
   * Filtre serveur par statut workflow (field_10 en SharePoint).
   * Valeurs possibles : 'Ouvert' / 'En cours' / 'Resolu' / 'Clos' / ''
   * (chaîne vide = pas de filtre = toutes les anomalies retournées).
   */
  const [filterStatut, setFilterStatut] = useState('')
  // Filtres client (champs personne + recherche libre, gérés côté frontend)
  const [filterAgent, setFilterAgent] = useState('')
  const [filterAffecte, setFilterAffecte] = useState('')
  const [filterDeclarant, setFilterDeclarant] = useState('')
  const [filterSearch, setFilterSearch] = useState('')
  // Snapshot des filtres client au moment du clic "Rechercher"
  const [appliedAgent, setAppliedAgent] = useState('')
  const [appliedAffecte, setAppliedAffecte] = useState('')
  const [appliedDeclarant, setAppliedDeclarant] = useState('')
  const [appliedSearch, setAppliedSearch] = useState('')


  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS — MODALES
   * Chaque modale a son propre groupe d'états. La présence d'un objet
   * non-null dans le state principal contrôle l'affichage de la modale.
   * ────────────────────────────────────────────────────────────────────── */

  /** Modale "Détail" (lecture complète d'une anomalie). null = fermée. */
  const [detailItem, setDetailItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)

  /**
   * Mode édition dans la modale Détail.
   *   - false : affichage en lecture seule (<dl>/<dd>)
   *   - true  : tous les champs métier deviennent des inputs/selects
   *
   * Le toggle se fait via le bouton "Modifier" en header (si l'utilisateur a
   * les droits — cf. canEditDetail). Quitter la modale (close ou changement
   * d'item) repasse implicitement en lecture seule via closeDetailModal().
   *
   * Les champs personnes (auteur, déclarant, personne affectée) NE SONT PAS
   * éditables ici : la personne affectée se change via le bouton "Affecter"
   * dédié (workflow personne + délai + commentaire en un seul geste).
   */
  const [detailEditMode, setDetailEditMode] = useState(false)

  /**
   * Snapshot des champs en cours d'édition dans la modale Détail.
   *
   * Initialisé depuis detailItem au passage en mode édition (enterDetailEdit).
   * Les inputs sont liés à ce state — l'item SharePoint n'est mis à jour
   * QU'AU clic sur "Enregistrer" (saveDetailEdit).
   *
   * Si l'utilisateur clique "Annuler", on jette ce snapshot sans rien écrire.
   */
  const [detailForm, setDetailForm] = useState({
    field_0: '',        // Date de l'anomalie (YYYY-MM-DD)
    field_4: '',        // Cause / description (texte brut, sans HTML)
    field_5: '',        // Classification
    field_6: '',        // Agence (ID en string)
    field_7: '',        // Réseau (ID en string, auto-déduit de l'agence)
    field_8: 0,         // Montant
    field_9: '',        // Date de régularisation (YYYY-MM-DD)
    field_10: '',       // Statut workflow
    criticiteAnomalie: '',  // Criticité
    domaineActivite: '',    // Domaine d'activité (référentiel DOMAINE_ACTIVITE_OPTIONS)
    delai: '',          // Délai de traitement en jours (entier sous forme string)
    commentaireAffectation: '',  // Commentaire pour la personne affectée
  })
  const [detailSaving, setDetailSaving] = useState(false)
  const [detailError, setDetailError] = useState<string | null>(null)

  /**
   * Modale "Affecter" — workflow en 2 étapes :
   *   1. Rechercher un utilisateur Office 365 et le sélectionner
   *   2. Renseigner le délai de traitement + le commentaire d'affectation
   *   3. Cliquer sur "Valider l'affectation" → écrit personneAffecter, delai
   *      et commentaireAffectation sur l'item SharePoint
   *
   * Champs SharePoint cibles :
   *   - personneAffecter      : Person (format Claims)
   *   - delai                 : texte (date YYYY-MM-DD stockée comme string)
   *   - commentaireAffectation: texte libre
   */
  const [affectItemId, setAffectItemId] = useState<number | null>(null)
  const [affectSearch, setAffectSearch] = useState('')
  const [affectResults, setAffectResults] = useState<User[]>([])
  const [affectLoading, setAffectLoading] = useState(false)
  /** Utilisateur sélectionné mais pas encore validé (étape 2 du wizard). */
  const [affectSelectedUser, setAffectSelectedUser] = useState<User | null>(null)
  /**
   * Délai de traitement EN NOMBRE DE JOURS (entier positif, stocké en string SP).
   * Ex: '7', '15', '30'. SharePoint a typé ce champ comme string max 255, donc
   * on stocke un nombre converti en string.
   */
  // Délai par défaut à l'ouverture de la modale d'affectation : 3 jours.
  // L'utilisateur peut l'ajuster avant validation.
  const [affectDelai, setAffectDelai] = useState('3')
  /** Commentaire d'affectation libre (instructions pour la personne assignée). */
  const [affectCommentaire, setAffectCommentaire] = useState('')
  /** Message d'erreur de validation (champs vides, etc.). */
  const [affectError, setAffectError] = useState<string | null>(null)

  /** Modale "Changer le statut" — workflow rapide (Ouvert/En cours/Resolu/Clos). */
  const [statusItem, setStatusItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)
  const [statusValue, setStatusValue] = useState('')
  const [statusDateRegul, setStatusDateRegul] = useState('')
  const [statusDateCloture, setStatusDateCloture] = useState('')
  const [statusSaving, setStatusSaving] = useState(false)
  const [statusError, setStatusError] = useState<string | null>(null)

  /** Modale "Suivi du ticket" — vue synthétique avec actions. */
  const [ticketItem, setTicketItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)

  /**
   * Modale "Clôture de la résolution" — workflow complet contrôleur :
   *   - Statut final, dates, auteur (modifiable)
   *   - Cause immédiate, cause racine, actions menées, observations
   *   - Pièce jointe (preuve de résolution, ajoutée à urlPieceJointe)
   */
  const [resolutionItem, setResolutionItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)
  const [resolutionForm, setResolutionForm] = useState({
    statut: 'Resolu',
    dateRegul: new Date().toISOString().split('T')[0],
    dateCloture: new Date().toISOString().split('T')[0],
    causeImmediate: '',
    causeRacine: '',
    actionsMenees: '',
    observations: '',
    /**
     * Mode de traitement décidé à la clôture. Référentiel
     * TYPE_SANCTION_OPTIONS. Champ propre à la clôture — pas saisi à la
     * création. Optionnel (un blâme/avertissement n'est pas systématique).
     */
    typeSanction: '',
  })
  const [resolutionSaving, setResolutionSaving] = useState(false)
  const [resolutionError, setResolutionError] = useState<string | null>(null)
  const [resolutionAttachment, setResolutionAttachment] = useState<File | null>(null)

  // Auteur modifiable depuis la modale clôture (autocomplétion Office 365)
  const [resolutionAuteurSearch, setResolutionAuteurSearch] = useState('')
  const [resolutionAuteurEmail, setResolutionAuteurEmail] = useState('')
  const [resolutionAuteurResults, setResolutionAuteurResults] = useState<User[]>([])
  const [showResolutionAuteurDropdown, setShowResolutionAuteurDropdown] = useState(false)


  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS — UI / RÉFÉRENTIELS
   * ────────────────────────────────────────────────────────────────────── */

  /** Toggle d'expansion des colonnes intermédiaires du tableau. */
  const [expandedColumns, setExpandedColumns] = useState(false)

  /** Référentiels (chargés une fois au montage, utilisés pour les selects/labels). */
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])


  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS — AUTOCOMPLÉTIONS DU FORMULAIRE DE CRÉATION
   * Champs auteur + personne affectée avec recherche Office 365 en temps réel.
   * ────────────────────────────────────────────────────────────────────── */

  // Auteur (formulaire création)
  const [auteurSearch, setAuteurSearch] = useState('')
  const [auteurEmail, setAuteurEmail] = useState('')
  const [auteurResults, setAuteurResults] = useState<User[]>([])
  const [showAuteurDropdown, setShowAuteurDropdown] = useState(false)

  // Personne affectée (formulaire création)
  //
  // ⚠ Pré-remplissage : par défaut, la personne affectée est le déclarant
  // lui-même (userName / userEmail). Le déclarant peut évidemment modifier
  // ce choix pour affecter à un collègue. Rationale : dans le workflow
  // habituel, le contrôleur qui déclare est aussi celui qui va traiter —
  // on évite ainsi une étape d'affectation systématique. Le manager peut
  // toujours réaffecter ensuite via le bouton "Affecter" du tableau.
  const [affecteFormSearch, setAffecteFormSearch] = useState(() => userName ?? '')
  const [affecteFormEmail, setAffecteFormEmail] = useState(() => userEmail ?? '')
  const [affecteFormResults, setAffecteFormResults] = useState<User[]>([])
  const [showAffecteFormDropdown, setShowAffecteFormDropdown] = useState(false)

  // Re-synchronise le pré-remplissage si l'email/nom du user changent
  // pendant la vie du composant (rare mais possible via bascule de rôle
  // simulé). Le useEffect ne s'exécute que si le champ est resté sur le
  // choix par défaut — on ne veut pas écraser une saisie manuelle.
  useEffect(() => {
    setAffecteFormEmail(prev => (prev === '' || prev === userEmail) ? (userEmail ?? '') : prev)
    setAffecteFormSearch(prev => (prev === '' || prev === userName) ? (userName ?? '') : prev)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userEmail, userName])


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — RECHERCHE D'UTILISATEURS (AUTOCOMPLÉTIONS)
   *
   * Pattern commun :
   *   - Met à jour le texte de recherche
   *   - Si < 2 caractères : reset les résultats, ferme la dropdown
   *   - Sinon : appel SearchUser(term, 10) puis affichage
   *   - Sélection d'un résultat : remplit l'email, ferme la dropdown
   * ────────────────────────────────────────────────────────────────────── */

  /** Recherche pour le champ "Auteur" du formulaire de création. */
  const searchAuteur = async (term: string) => {
    setAuteurSearch(term)
    if (term.length < 2) { setAuteurResults([]); setShowAuteurDropdown(false); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) { setAuteurResults(result.data); setShowAuteurDropdown(true) }
    } catch (err) { console.error('Erreur recherche auteur', err) }
  }

  const selectAuteur = (u: User) => {
    setAuteurEmail(u.Mail ?? '')
    setAuteurSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAuteurDropdown(false)
    setAuteurResults([])
  }

  const searchAffecteForm = async (term: string) => {
    setAffecteFormSearch(term)
    if (term.length < 2) { setAffecteFormResults([]); setShowAffecteFormDropdown(false); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) { setAffecteFormResults(result.data); setShowAffecteFormDropdown(true) }
    } catch (err) { console.error('Erreur recherche personne affectee', err) }
  }

  const selectAffecteForm = (u: User) => {
    setAffecteFormEmail(u.Mail ?? '')
    setAffecteFormSearch(u.DisplayName ?? u.Mail ?? '')
    setShowAffecteFormDropdown(false)
    setAffecteFormResults([])
  }


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — MODALE D'AFFECTATION
   *
   * Workflow : ouvrir la modale → rechercher un utilisateur → sélectionner
   * → SharePoint update du champ personneAffecter avec format Claims.
   * ────────────────────────────────────────────────────────────────────── */

  /** Recherche pour la modale "Affecter" — utilise affectSearch dédié. */
  const searchAffectUser = async (term: string) => {
    setAffectSearch(term)
    if (term.length < 2) { setAffectResults([]); return }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) setAffectResults(result.data)
    } catch (err) { console.error('Erreur recherche affectation', err) }
  }

  /**
   * Étape 1 → 2 : un utilisateur a été cliqué dans la liste de résultats.
   *
   * On NE confirme PAS encore l'affectation à SharePoint : on bascule la modale
   * en mode "formulaire" (délai + commentaire). La validation effective passe
   * désormais par le clic sur le bouton "Valider l'affectation" (cf. confirmAffect).
   *
   * On vide aussi les résultats de recherche pour laisser place aux champs
   * du formulaire, mais on conserve le terme recherché (visuellement pratique :
   * l'utilisateur voit ce qu'il a tapé).
   */
  const selectAffectUser = (u: User) => {
    setAffectSelectedUser(u)
    setAffectResults([])
    setAffectError(null)
  }

  /**
   * Permet de revenir à l'étape 1 (changer de personne sans fermer la modale).
   * Vide la sélection et remet le focus sur la recherche.
   */
  const clearAffectSelection = () => {
    setAffectSelectedUser(null)
    setAffectError(null)
  }

  /**
   * Confirme l'affectation : met à jour personneAffecter + delai +
   * commentaireAffectation sur l'item SharePoint.
   *
   * Format SharePoint personne :
   *   {
   *     '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
   *     Claims: 'i:0#.f|membership|email@domaine.com'
   *   }
   *
   * Le `as never` cast contourne le typage trop strict du service généré
   * (le modèle attend personneAffecterValue complet, mais SharePoint accepte
   * juste les Claims pour les écritures).
   *
   * Règle métier : délai et commentaire sont REQUIS pour valider une affectation
   * (le métier veut tracer qui a été assigné, pour quand, avec quelles
   * instructions). Garde-fou ici + désactivation visuelle du bouton dans l'UI.
   */
  const confirmAffect = async () => {
    if (!affectItemId || !affectSelectedUser?.Mail) {
      setAffectError('Sélectionnez d\'abord une personne.')
      return
    }
    // Validation du délai : doit être un entier strictement positif.
    // On le parse pour rejeter "0", "-3", "abc", "" en une seule expression.
    const delaiNum = parseInt(affectDelai, 10)
    if (!Number.isFinite(delaiNum) || delaiNum <= 0) {
      setAffectError('Le délai de traitement doit être un nombre de jours > 0.')
      return
    }
    if (!affectCommentaire.trim()) {
      setAffectError('Le commentaire d\'affectation est obligatoire.')
      return
    }
    setAffectLoading(true)
    setAffectError(null)
    try {
      await DCPO_LISTE_ANORMALIEService.update(String(affectItemId), {
        personneAffecter: {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(affectSelectedUser.Mail),
        } as never,
        // delai : stocké comme string (typage SP) mais représente un entier
        // de jours. On envoie la version normalisée (sans espaces / zéros tête).
        delai: String(delaiNum),
        commentaireAffectation: affectCommentaire.trim(),
        // dateAffection : horodatage du moment de l'affectation. Sert de
        // référence pour calculer l'échéance affichée dans le tableau
        // (cf. colonne "Échéance / délai" et helper getDeadlineStatus).
        //
        // NB : le champ est typé `dateAffection` côté SP (typo "Affection"
        // sans 't' — conservé tel quel pour matcher la colonne réelle).
        dateAffection: new Date().toISOString(),
      })

      // Notification Teams (fire-and-forget). Récupère le titre de
      // l'anomalie depuis le state local pour donner un contexte clair
      // au destinataire (numéro de ticket + titre).
      const affectedItem = items.find(it => it.ID === affectItemId)
      const ticketNum = affectItemId ? `T-${affectItemId}` : 'Anomalie'
      const titre = affectedItem?.Title?.trim() || ''
      notifyAffectation({
        type: 'anomalie',
        email: affectSelectedUser.Mail,
        subject: titre ? `${ticketNum} — ${titre}` : ticketNum,
        details: `Délai de traitement : ${delaiNum} jour(s).${affectCommentaire.trim() ? ` Commentaire : ${affectCommentaire.trim()}` : ''}`,
      })

      closeAffectModal()
      await fetchItems()
    } catch (err) {
      console.error('Erreur affectation', err)
      setAffectError('Échec de l\'affectation. Réessayer.')
    }
    finally { setAffectLoading(false) }
  }

  /** Reset complet de la modale affectation à la fermeture. */
  const closeAffectModal = () => {
    setAffectItemId(null)
    setAffectSearch('')
    setAffectResults([])
    setAffectSelectedUser(null)
    // Délai remis à 3 jours (valeur par défaut métier — cohérent avec
    // l'init du useState ci-dessus).
    setAffectDelai('3')
    setAffectCommentaire('')
    setAffectError(null)
  }


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — MODE ÉDITION DANS LA MODALE DÉTAIL
   *
   * Workflow :
   *   1. Clic "Modifier" → enterDetailEdit() initialise detailForm depuis
   *      detailItem et bascule en mode édition
   *   2. L'utilisateur saisit dans les inputs (detailForm est mis à jour)
   *   3. Clic "Annuler" → cancelDetailEdit() jette les modifications
   *      OU Clic "Enregistrer" → saveDetailEdit() écrit en SP
   *
   * La modale Détail elle-même est fermée séparément via closeDetailModal()
   * qui prend soin de reset le mode édition (évite la conservation d'un
   * snapshot orphelin).
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Helper : ferme la modale Détail et reset l'éventuel mode édition en cours.
   *
   * À utiliser à la place de setDetailItem(null) pour garantir la cohérence
   * des états (sinon le snapshot detailForm reste en mémoire jusqu'au prochain
   * enterDetailEdit, ce qui peut polluer une future session d'édition).
   */
  const closeDetailModal = () => {
    setDetailItem(null)
    setDetailEditMode(false)
    setDetailError(null)
    setDetailSaving(false)
  }

  /**
   * Bascule en mode édition : initialise detailForm depuis detailItem courant.
   *
   * Pourquoi un snapshot plutôt que pointer directement detailItem ?
   *   - detailItem est partagé avec la liste : modifier ses champs directement
   *     polluerait l'affichage de la liste pendant l'édition
   *   - permet le "Annuler" propre (jeter le snapshot, garder detailItem intact)
   *
   * Le champ field_4 (cause) contient du HTML en SharePoint ; on le convertit
   * en texte brut via stripHtml pour l'édition (la sauvegarde renvoie du
   * texte brut, SP l'accepte tel quel — il ne sera juste pas formaté).
   */
  const enterDetailEdit = () => {
    if (!detailItem) return
    setDetailForm({
      field_0: detailItem.field_0 ? detailItem.field_0.split('T')[0] : '',
      field_4: detailItem.field_4 ? stripHtml(detailItem.field_4) : '',
      field_5: detailItem.field_5 ?? '',
      field_6: detailItem.field_6 ?? '',
      field_7: detailItem.field_7 ?? '',
      field_8: detailItem.field_8 ?? 0,
      field_9: detailItem.field_9 ? detailItem.field_9.split('T')[0] : '',
      field_10: detailItem.field_10 ?? '',
      criticiteAnomalie: detailItem.criticiteAnomalie ?? '',
      domaineActivite: detailItem.domaineActivite ?? '',
      delai: detailItem.delai ?? '',
      commentaireAffectation: detailItem.commentaireAffectation ?? '',
    })
    setDetailEditMode(true)
    setDetailError(null)
  }

  /** Sort du mode édition SANS sauvegarder. */
  const cancelDetailEdit = () => {
    setDetailEditMode(false)
    setDetailError(null)
  }

  /** Helper générique pour patcher un champ du formulaire d'édition. */
  const updateDetailForm = <K extends keyof typeof detailForm>(key: K, value: (typeof detailForm)[K]) => {
    setDetailForm(prev => ({ ...prev, [key]: value }))
  }

  /**
   * Envoie les modifications en SharePoint puis recharge la liste.
   *
   * Validation minimale :
   *   - Si délai renseigné, doit être un entier > 0 (cohérent avec le workflow
   *     d'affectation)
   *   - Les autres champs n'ont pas de validation stricte (on accepte vide
   *     pour permettre l'effacement)
   *
   * Après succès :
   *   - On recharge la liste via fetchItems()
   *   - On bascule la modale en lecture seule (mais on garde detailItem ouvert
   *     pour que l'utilisateur voie les nouvelles valeurs)
   *   - Le nouveau detailItem est récupéré dans le résultat du refetch via
   *     un match sur l'ID
   */
  const saveDetailEdit = async () => {
    if (!detailItem?.ID) return

    // Validation conditionnelle du délai
    if (detailForm.delai.trim() !== '') {
      const delaiNum = parseInt(detailForm.delai, 10)
      if (!Number.isFinite(delaiNum) || delaiNum <= 0) {
        setDetailError('Le délai doit être un nombre de jours > 0.')
        return
      }
    }

    setDetailSaving(true)
    setDetailError(null)
    try {
      const payload: Record<string, unknown> = {
        field_0: detailForm.field_0 ? `${detailForm.field_0}T00:00:00Z` : '',
        field_4: detailForm.field_4,
        field_5: detailForm.field_5,
        field_6: detailForm.field_6,
        field_7: detailForm.field_7,
        field_8: Number(detailForm.field_8) || 0,
        field_9: detailForm.field_9 ? `${detailForm.field_9}T00:00:00Z` : '',
        field_10: detailForm.field_10,
        criticiteAnomalie: detailForm.criticiteAnomalie,
        domaineActivite: detailForm.domaineActivite,
        // Normalisation du délai : on stocke l'entier sans zéros tête / espaces
        delai: detailForm.delai.trim() === '' ? '' : String(parseInt(detailForm.delai, 10)),
        commentaireAffectation: detailForm.commentaireAffectation.trim(),
      }
      await DCPO_LISTE_ANORMALIEService.update(String(detailItem.ID), payload as never)
      await fetchItems()
      // On garde la modale ouverte mais on revient en lecture seule.
      // detailItem sera rafraîchi automatiquement au prochain render via fetchItems.
      setDetailEditMode(false)
      // Mise à jour optimiste de detailItem pour refléter les changements
      // immédiatement (sinon il faut attendre que le user reclique sur Détail).
      setDetailItem(prev => prev ? { ...prev, ...payload } as DCPO_LISTE_ANORMALIERead : prev)
    } catch (err) {
      console.error('Erreur sauvegarde anomalie', err)
      setDetailError('Échec de la sauvegarde. Réessayer.')
    } finally {
      setDetailSaving(false)
    }
  }


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — MODALE CHANGEMENT DE STATUT (workflow rapide)
   *
   * Permet de basculer le statut d'une anomalie sans passer par le formulaire
   * complet de résolution. Use case : Ouvert → En cours, ou réouverture
   * d'un Clos.
   *
   * Quand le statut devient Resolu ou Clos, on remplit AUSSI :
   *   - field_9 (date de régularisation)
   *   - date_cloture_ticket (date de clôture)
   * ────────────────────────────────────────────────────────────────────── */

  /** Ouvre la modale en pré-remplissant avec les valeurs courantes de l'item. */
  const openStatusModal = (item: DCPO_LISTE_ANORMALIERead) => {
    setStatusItem(item)
    setStatusValue(item.field_10 ?? '')
    const today = new Date().toISOString().split('T')[0]
    setStatusDateRegul(item.field_9 ? item.field_9.split('T')[0] : today)
    setStatusDateCloture(item.date_cloture_ticket ? item.date_cloture_ticket.split('T')[0] : today)
    setStatusError(null)
  }

  const closeStatusModal = () => {
    setStatusItem(null)
    setStatusValue('')
    setStatusDateRegul('')
    setStatusDateCloture('')
    setStatusError(null)
    setStatusSaving(false)
  }

  const saveStatus = async () => {
    if (!statusItem || !statusValue) {
      setStatusError('Veuillez choisir un statut.')
      return
    }
    const closesTicket = statusValue === 'Resolu' || statusValue === 'Clos'
    // Règle métier : la date de clôture ne peut pas être antérieure à la
    // date d'ouverture du ticket (cf. même règle dans saveResolution).
    if (closesTicket) {
      const openingDateOnly = statusItem.dateOuvertureTicket
        ? statusItem.dateOuvertureTicket.split('T')[0]
        : ''
      if (openingDateOnly && statusDateCloture && statusDateCloture < openingDateOnly) {
        setStatusError("La date de clôture ne peut pas être antérieure à la date d'ouverture du ticket.")
        return
      }
    }
    setStatusSaving(true)
    setStatusError(null)
    try {
      const payload: Record<string, unknown> = { field_10: statusValue }
      const today = new Date().toISOString().split('T')[0]
      if (closesTicket) {
        const regulDate = statusDateRegul || today
        payload.field_9 = `${regulDate}T00:00:00Z`
        // Le ticket est clos dès qu'il passe à Resolu ou Clos
        const clotureDate = statusDateCloture || today
        payload.date_cloture_ticket = `${clotureDate}T00:00:00Z`
      }
      const result = await DCPO_LISTE_ANORMALIEService.update(
        String(statusItem.ID),
        payload as Partial<Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>>,
      )
      if (!result.success) {
        setStatusError('Erreur lors de la mise à jour du statut.')
        return
      }
      closeStatusModal()
      await fetchItems()
    } catch (err) {
      console.error('Erreur changement statut', err)
      setStatusError('Erreur lors de la mise à jour du statut.')
    } finally {
      setStatusSaving(false)
    }
  }


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — MODALE CLÔTURE DE LA RÉSOLUTION (workflow complet contrôleur)
   *
   * Workflow le plus riche : le contrôleur consigne l'ensemble des éléments
   * de la résolution (statut, dates, auteur, causes, actions, observations,
   * preuve attachée). Les données sont enregistrées dans :
   *   - field_10 (statut)
   *   - field_9 (date régularisation)
   *   - date_cloture_ticket
   *   - field_4 (HTML structuré : description riche du contrôleur)
   *   - urlPieceJointe (pièce jointe ajoutée à la liste existante)
   *   - auteur_anormalie (si modifié)
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Ouvre la modale de clôture en pré-remplissant tous les champs avec
   * les valeurs actuelles de l'item (édition possible).
   */
  const openResolution = (item: DCPO_LISTE_ANORMALIERead) => {
    const today = new Date().toISOString().split('T')[0]
    const existingDescription = item.field_4 ? stripHtml(item.field_4) : ''
    setResolutionForm({
      statut: item.field_10 === 'Clos' ? 'Clos' : 'Resolu',
      dateRegul: item.field_9 ? item.field_9.split('T')[0] : today,
      dateCloture: item.date_cloture_ticket ? item.date_cloture_ticket.split('T')[0] : today,
      causeImmediate: existingDescription,
      causeRacine: '',
      actionsMenees: '',
      observations: '',
      // Pré-remplit avec la valeur déjà stockée (cas de réouverture d'une
      // clôture précédente), sinon laisse vide pour forcer une saisie consciente.
      typeSanction: item.typeSanction ?? '',
    })
    setResolutionAttachment(null)
    setResolutionError(null)
    // Pré-remplir l'auteur avec celui actuellement enregistré
    setResolutionAuteurEmail(item.auteur_anormalie?.Email ?? '')
    setResolutionAuteurSearch(item.auteur_anormalie?.DisplayName ?? item.auteur_anormalie?.Email ?? '')
    setResolutionAuteurResults([])
    setShowResolutionAuteurDropdown(false)
    setResolutionItem(item)
    setTicketItem(null)
  }

  const closeResolution = () => {
    setResolutionItem(null)
    setResolutionError(null)
    setResolutionSaving(false)
    setResolutionAttachment(null)
    setResolutionAuteurSearch('')
    setResolutionAuteurEmail('')
    setResolutionAuteurResults([])
    setShowResolutionAuteurDropdown(false)
  }

  const searchResolutionAuteur = async (term: string) => {
    setResolutionAuteurSearch(term)
    if (term.length < 2) {
      setResolutionAuteurResults([])
      setShowResolutionAuteurDropdown(false)
      return
    }
    try {
      const result = await Office365UsersService.SearchUser(term, 10)
      if (result.data) {
        setResolutionAuteurResults(result.data)
        setShowResolutionAuteurDropdown(true)
      }
    } catch (err) {
      console.error('Erreur recherche auteur résolution', err)
    }
  }

  const selectResolutionAuteur = (u: User) => {
    setResolutionAuteurEmail(u.Mail ?? '')
    setResolutionAuteurSearch(u.DisplayName ?? u.Mail ?? '')
    setShowResolutionAuteurDropdown(false)
    setResolutionAuteurResults([])
  }

  const updateResolutionForm = <K extends keyof typeof resolutionForm>(field: K, value: string) => {
    setResolutionForm(prev => ({ ...prev, [field]: value }))
  }

  const buildResolutionHtml = (): string => {
    const blocks: string[] = []
    if (resolutionForm.causeImmediate.trim()) {
      blocks.push(`<p><strong>Cause immédiate :</strong> ${escapeHtml(resolutionForm.causeImmediate)}</p>`)
    }
    if (resolutionForm.causeRacine.trim()) {
      blocks.push(`<p><strong>Cause racine :</strong> ${escapeHtml(resolutionForm.causeRacine)}</p>`)
    }
    // Les actions menées NE SONT PLUS écrites ici (concaténation dans field_4).
    // Elles sont désormais stockées dans la colonne dédiée `actionsMenees` de
    // la liste SP (cf. payload de saveResolution).
    if (resolutionForm.observations.trim()) {
      blocks.push(`<p><strong>Observations :</strong> ${escapeHtml(resolutionForm.observations)}</p>`)
    }
    return blocks.join('')
  }

  const saveResolution = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!resolutionItem) return
    if (!resolutionForm.actionsMenees.trim()) {
      setResolutionError('Les actions menées sont obligatoires pour clôturer une résolution.')
      return
    }
    // Règle métier : la date de clôture ne peut pas être antérieure à la
    // date d'ouverture du ticket (dateOuvertureTicket, horodatée à la
    // création — cf. handleSubmit). Comparaison en string 'YYYY-MM-DD' :
    // le format ISO trie lexicographiquement comme chronologiquement.
    const openingDateOnly = resolutionItem.dateOuvertureTicket
      ? resolutionItem.dateOuvertureTicket.split('T')[0]
      : ''
    if (openingDateOnly && resolutionForm.dateCloture && resolutionForm.dateCloture < openingDateOnly) {
      setResolutionError("La date de clôture ne peut pas être antérieure à la date d'ouverture du ticket.")
      return
    }
    setResolutionSaving(true)
    setResolutionError(null)
    try {
      const today = new Date().toISOString().split('T')[0]
      const html = buildResolutionHtml()
      // La modale "Clore la résolution" ferme toujours le ticket,
      // que le statut final soit Resolu ou Clos.
      const payload: Record<string, unknown> = {
        field_10: resolutionForm.statut,
        field_9: `${resolutionForm.dateRegul || today}T00:00:00Z`,
        date_cloture_ticket: `${resolutionForm.dateCloture || today}T00:00:00Z`,
        // typeSanction n'est écrit QU'À LA CLÔTURE — c'est la seule
        // entrée utilisateur pour ce champ dans l'application.
        typeSanction: resolutionForm.typeSanction,
        // actionsMenees : colonne SP dédiée (ajoutée récemment). Auparavant
        // ces actions étaient concaténées dans field_4 ; désormais on les
        // écrit ici directement (les éventuels retours à la ligne du
        // textarea sont conservés tels quels — c'est SP qui décide du
        // formatage à l'affichage natif).
        actionsMenees: resolutionForm.actionsMenees.trim(),
      }
      if (html) payload.field_4 = html
      // Mise à jour de l'auteur si modifié
      const previousAuteurEmail = resolutionItem.auteur_anormalie?.Email ?? ''
      if (resolutionAuteurEmail && resolutionAuteurEmail !== previousAuteurEmail) {
        payload.auteur_anormalie = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(resolutionAuteurEmail),
        }
      }
      const result = await DCPO_LISTE_ANORMALIEService.update(
        String(resolutionItem.ID),
        payload as Partial<Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>>,
      )
      if (!result.success) {
        setResolutionError('Erreur lors de l\'enregistrement de la résolution.')
        return
      }

      // Notification Teams à l'AUTEUR si son email a changé lors de la clôture.
      // Court-circuit défensif : on ne notifie pas si l'auteur est aussi le
      // manager qui fait la clôture (inutile de s'auto-notifier).
      if (
        resolutionAuteurEmail
        && resolutionAuteurEmail !== previousAuteurEmail
        && resolutionAuteurEmail.toLowerCase() !== (userEmail ?? '').toLowerCase()
      ) {
        const description = resolutionItem.field_4
          ? stripHtml(resolutionItem.field_4).slice(0, 80)
          : ''
        notifyAuteurAnomalie({
          email: resolutionAuteurEmail,
          anomalieLabel: `T-${resolutionItem.ID}${description ? ' — ' + description : ''}`,
          declareParName: userName,
        })
      }

      // Upload de la pièce jointe (preuve de résolution)
      //
      // La concaténation via appendUrl produit une valeur COMPATIBLE avec
      // le format `uri` du champ SP :
      //   - 1 seule URL → l'URL brute (cliquable dans l'UI SP native)
      //   - N URLs      → encapsulées dans une data URI base64
      //                   (`data:text/x-dcpo-urls;base64,...`) qui reste
      //                   une URI RFC 2397 valide.
      // Voir la SECTION 6 de ticketAttachments.ts pour le détail du format.
      //
      // Cela permet de conserver l'URL saisie à la déclaration ET d'y
      // ajouter la nouvelle preuve de résolution sans écraser l'historique.
      if (resolutionAttachment) {
        try {
          const attachmentUrl = await uploadAttachment(String(resolutionItem.ID), resolutionAttachment)
          if (attachmentUrl) {
            const concatenated = appendUrl(resolutionItem.urlPieceJointe, attachmentUrl)
            await DCPO_LISTE_ANORMALIEService.update(String(resolutionItem.ID), {
              urlPieceJointe: concatenated,
            })
          }
        } catch (err) {
          console.error('Erreur upload pièce jointe résolution:', err)
          setResolutionError('Statut enregistré mais l\'upload de la pièce jointe a échoué.')
          await fetchItems()
          return
        }
      }

      closeResolution()
      await fetchItems()
    } catch (err) {
      console.error('Erreur résolution', err)
      setResolutionError('Erreur lors de l\'enregistrement de la résolution.')
    } finally {
      setResolutionSaving(false)
    }
  }


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — CHARGEMENT, FILTRES, RECHERCHE
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Construit la clause $filter OData à envoyer à SharePoint.
   *
   * Format : "field eq 'valeur' and field2 eq 'valeur2'"
   *
   * Champs utilisés :
   *   - Created (auto-indexé SP) pour les filtres de date
   *   - field_5, field_6, field_7, field_10, criticiteAnomalie
   *
   * Les filtres "Agent" et "Personne affectée" NE sont PAS dans cette clause
   * (champs Personne difficiles à filtrer en OData via le connecteur) →
   * filtrage côté client via filteredItems plus bas.
   *
   * IMPORTANT — performance : les colonnes utilisées en $filter doivent être
   * indexées dans SharePoint pour rester performantes au-delà de 5 000 items
   * (List View Threshold). Voir la doc projet sur l'indexation.
   */
  const buildFilter = () => {
    const clauses: string[] = []
    if (filterDateFrom) {
      clauses.push(`Created ge '${filterDateFrom}T00:00:00Z'`)
    }
    if (filterDateTo) {
      clauses.push(`Created le '${filterDateTo}T23:59:59Z'`)
    }
    if (filterAgence) clauses.push(`field_6 eq '${filterAgence}'`)
    if (filterReseau) clauses.push(`field_7 eq '${filterReseau}'`)
    if (filterClassification) clauses.push(`field_5 eq '${filterClassification}'`)
    if (filterCriticite) clauses.push(`criticiteAnomalie eq '${filterCriticite}'`)
    if (filterDomaineActivite) clauses.push(`domaineActivite eq '${filterDomaineActivite}'`)
    if (filterStatut) {
      clauses.push(`field_10 eq '${filterStatut}'`)
    } else {
      // Règle métier : la liste principale des anomalies N'AFFICHE PAS les
      // anomalies clôturées ou résolues — elles sont consultables sur la
      // page "Fiche récapitulatif de l'anomalie". On n'applique l'exclusion
      // QUE si l'utilisateur n'a pas choisi explicitement un statut.
      clauses.push("field_10 ne 'Clos'")
      clauses.push("field_10 ne 'Resolu'")
    }
    /**
     * Pas de restriction de visibilité par rôle : tous les utilisateurs
     * (Controleur compris) voient l'ensemble des anomalies en cours, y
     * compris celles affectées à d'autres contrôleurs. Cela facilite la
     * coordination intra-équipe (vision partagée du backlog DCPO).
     */
    return clauses.join(' and ')
  }

  /**
   * Récupère les anomalies depuis SharePoint avec les filtres serveur courants.
   *
   * Tri par défaut : Created desc → les plus récentes en premier (Created
   * étant auto-indexé, ce orderBy est performant même sur grosse liste).
   *
   * Le SDK Power Apps gère la pagination automatique via getAll().
   */
  const fetchItems = async () => {
    setLoading(true)
    try {
      const filter = buildFilter()
      const options: { orderBy: string[]; filter?: string } = {
        orderBy: ['Created desc'],
      }
      if (filter) options.filter = filter

      const items = await getAllPages<DCPO_LISTE_ANORMALIERead>(
        DCPO_LISTE_ANORMALIEService,
        options,
      )
      setItems(items)
    } catch (err) { console.error('Erreur chargement anomalies', err) }
    finally { setLoading(false) }
  }

  useEffect(() => {
    const loadLists = async () => {
      try {
        const [agencesRows, reseauxRows] = await Promise.all([
          loadAgences(),
          loadReseaux(),
        ])
        setAgences(agencesRows)
        setReseaux(reseauxRows)
      } catch (err) { console.error('Erreur chargement agences/reseaux', err) }
    }
    loadLists()
    fetchItems()
    // userRole/userEmail intégrés : si l'utilisateur simule un autre rôle via
    // le sélecteur de la topbar, on doit refetch pour appliquer (ou retirer)
    // la restriction "Controleur ne voit que ses anomalies".
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userRole, userEmail])

  /**
   * Au clic "Rechercher" :
   *   1. Snapshot des filtres client (agent/affecté) → applied*
   *   2. Refetch serveur avec les nouveaux filtres OData
   *
   * Découplage saisie / exécution : tant qu'on tape dans les inputs,
   * rien ne se passe. C'est le clic ICI qui déclenche tout.
   */
  const handleSearch = () => {
    setAppliedAgent(filterAgent)
    setAppliedAffecte(filterAffecte)
    setAppliedDeclarant(filterDeclarant)
    setAppliedSearch(filterSearch)
    fetchItems()
  }

  /**
   * Réinitialise tous les filtres (saisis ET appliqués) et refetch.
   *
   * setTimeout(0) : permet à tous les setState de s'appliquer avant le
   * refetch, évite un état intermédiaire avec d'anciens filtres encore actifs.
   */
  const handleResetFilters = () => {
    setFilterDateFrom('')
    setFilterDateTo('')
    setFilterAgence('')
    setFilterReseau('')
    setFilterClassification('')
    setFilterCriticite('')
    setFilterDomaineActivite('')
    setFilterStatut('')
    setFilterAgent('')
    setFilterAffecte('')
    setFilterDeclarant('')
    setFilterSearch('')
    setAppliedAgent('')
    setAppliedAffecte('')
    setAppliedDeclarant('')
    setAppliedSearch('')
    setTimeout(() => fetchItems(), 0)
  }

  /**
   * Filtrage CLIENT supplémentaire pour les champs personne (agent + affecté).
   *
   * Logique :
   *   1. Si aucun filtre appliqué → retourner items tel quel (perf)
   *   2. Sinon : filter() avec recherche partielle dans nom + email
   *
   * useMemo : évite de recalculer à chaque render si rien n'a changé.
   * Dépend de `items` (résultat serveur) + `appliedAgent` + `appliedAffecte`.
   */
  const filteredItems = useMemo(() => {
    const agentTerm = appliedAgent.trim().toLowerCase()
    const affecteTerm = appliedAffecte.trim().toLowerCase()
    const declarantTerm = appliedDeclarant.trim().toLowerCase()
    const searchTerm = appliedSearch.trim().toLowerCase()
    // Tous les utilisateurs (y compris les Controleurs) voient l'ensemble des
    // anomalies "en cours" — pas de restriction de visibilité par rôle ici.
    // Seuls les filtres saisis dans la barre (agent / personne affectée /
    // déclarant / recherche libre) peuvent restreindre la liste.
    if (!agentTerm && !affecteTerm && !declarantTerm && !searchTerm) return items
    return items.filter(it => {
      if (agentTerm) {
        const haystack = `${it.auteur_anormalie?.DisplayName ?? ''} ${it.auteur_anormalie?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(agentTerm)) return false
      }
      if (affecteTerm) {
        const haystack = `${it.personneAffecter?.DisplayName ?? ''} ${it.personneAffecter?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(affecteTerm)) return false
      }
      if (declarantTerm) {
        const haystack = `${it.declarant_anormalie?.DisplayName ?? ''} ${it.declarant_anormalie?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(declarantTerm)) return false
      }
      if (searchTerm) {
        const haystack = [
          it.ID ? `T-${it.ID}` : '',
          it.Title ?? '',
          it.field_3 ?? '',
          it.field_4 ?? '',
          it.field_5 ?? '',
          it.field_10 ?? '',
          it.criticiteAnomalie ?? '',
          it.domaineActivite ?? '',
          it.declarant_anormalie?.DisplayName ?? '',
          it.auteur_anormalie?.DisplayName ?? '',
          it.personneAffecter?.DisplayName ?? '',
        ].join(' ').toLowerCase()
        if (!haystack.includes(searchTerm)) return false
      }
      return true
    })
  }, [items, appliedAgent, appliedAffecte, appliedDeclarant, appliedSearch])

  /* ──────────────────────────────────────────────────────────────────────
   * PAGINATION
   *
   * Le tableau n'affiche que la page courante ; les stats cards restent
   * basées sur filteredItems (vue d'ensemble du filtre, pas de la page).
   * resetKey : remet à la page 1 quand les filtres ou les items changent.
   * ────────────────────────────────────────────────────────────────────── */
  const pagination = usePagination({
    total: filteredItems.length,
    resetKey: `${items.length}|${appliedAgent}|${appliedAffecte}|${appliedDeclarant}|${appliedSearch}`,
  })
  const pagedItems = useMemo(
    () => filteredItems.slice(pagination.start, pagination.end),
    [filteredItems, pagination.start, pagination.end],
  )

  // Apparition en cascade des lignes du tableau au (re)chargement des données
  const tableRef = useRef<HTMLDivElement>(null)
  useStagger(tableRef, { selector: 'tbody tr', deps: [pagedItems] })


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — FORMULAIRE DE CRÉATION
   * ────────────────────────────────────────────────────────────────────── */

  /** Met à jour un champ du formulaire (utilisé par les inputs). */
  const handleChange = (field: string, value: string | number) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  /** Reset complet du formulaire après création réussie ou annulation.
   *  Note : la personne affectée est remise au déclarant (défaut métier),
   *  pas à une chaîne vide — cohérent avec l'init dans useState. */
  const resetForm = () => {
    setForm(EMPTY_FORM)
    setAuteurSearch(''); setAuteurEmail('')
    setAffecteFormSearch(userName ?? ''); setAffecteFormEmail(userEmail ?? '')
    setAttachment(null)
  }

  /**
   * Upload une pièce jointe via Power Automate (workflow externe).
   * Identique à uploadTicketAttachment du lib mais avec params figés
   * pour la création (Choise = 'Visite').
   *
   * Note : duplication historique avec ticketAttachments.ts. Pourrait être
   * remplacée par un import du lib helper une fois testé.
   */
  const uploadAttachment = async (itemId: string, file: File): Promise<string | undefined> => {
    const fileContent = await fileToBase64(file)
    const response = await fetch(ATTACHMENT_API_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        fileName: file.name,
        fileContent,
        Choise: 'Visite',
        idItem: itemId,
      }),
    })
    if (!response.ok) {
      const errorText = await response.text()
      throw new Error(`Echec upload piece jointe (${response.status}): ${errorText}`)
    }
    const data: UploadResponse = await response.json()
    return data.attachmentUrl
  }

  /**
   * Soumission du formulaire de création d'anomalie.
   *
   * Étapes :
   *   1. Construire le payload SharePoint (uniquement les champs renseignés)
   *   2. Format date : ajouter T00:00:00Z (format ISO requis par SharePoint)
   *   3. Statut par défaut = 'Ouvert' si non choisi
   *   4. Ouvrir automatiquement le ticket : dateOuvertureTicket = maintenant
   *   5. Sérialiser les champs Personne (déclarant/auteur/affecté) au format Claims
   *   6. Créer l'item SharePoint
   *   7. Si pièce jointe : uploader puis update urlPieceJointe
   *   8. Reset du formulaire et refetch
   */
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setSubmitting(true)
    try {
      // Payload OData : on n'inclut un champ que s'il est renseigné
      // (évite d'écraser des valeurs avec des chaînes vides)
      const payload: Record<string, unknown> = {}
      if (form.field_0) payload.field_0 = form.field_0 + 'T00:00:00Z'
      if (form.field_4) payload.field_4 = form.field_4
      if (form.field_5) payload.field_5 = form.field_5
      if (form.field_6) payload.field_6 = form.field_6
      if (form.field_7) payload.field_7 = form.field_7
      if (form.field_8) payload.field_8 = form.field_8
      if (form.field_9) payload.field_9 = form.field_9 + 'T00:00:00Z'
      if (form.field_10) payload.field_10 = form.field_10
      else payload.field_10 = 'Ouvert'
      if (form.criticiteAnomalie) payload.criticiteAnomalie = form.criticiteAnomalie
      if (form.domaineActivite) payload.domaineActivite = form.domaineActivite

      // Ouverture automatique du ticket
      payload.dateOuvertureTicket = new Date().toISOString()

      if (auteurEmail) {
        payload['auteur_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(auteurEmail),
        }
      }
      if (userEmail) {
        payload['declarant_anormalie'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(userEmail),
        }
      }
      if (affecteFormEmail) {
        payload['personneAffecter'] = {
          '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
          Claims: toClaims(affecteFormEmail),
        }
        // Quand une personne est affectée dès la création, on horodate
        // l'affectation et on applique un délai par défaut de 3 jours.
        // L'utilisateur peut ensuite ajuster ce délai depuis la modale
        // d'affectation ou la modale d'édition du détail.
        payload['dateAffection'] = new Date().toISOString()
        payload['delai'] = '3'
      }

      const result = await DCPO_LISTE_ANORMALIEService.create(payload as Omit<DCPO_LISTE_ANORMALIEWrite, 'ID'>)
      if (!result.success) { console.error('Erreur SharePoint:', result.error); return }

      if (attachment && result.data?.ID) {
        try {
          const attachmentUrl = await uploadAttachment(String(result.data.ID), attachment)
          if (attachmentUrl) {
            await DCPO_LISTE_ANORMALIEService.update(String(result.data.ID), {
              urlPieceJointe: attachmentUrl,
            })
          }
        } catch (err) {
          console.error('Erreur upload piece jointe:', err)
        }
      }

      // Notification Teams si une personne a été affectée dès la création
      // (auquel cas elle vient de se voir attribuer un nouveau dossier).
      // Court-circuit : pas d'auto-notification si le déclarant s'affecte
      // lui-même (cas par défaut désormais courant — un contrôleur qui
      // s'attribue sa propre déclaration n'a pas besoin d'un Teams).
      if (
        affecteFormEmail
        && result.data?.ID
        && affecteFormEmail.toLowerCase() !== (userEmail ?? '').toLowerCase()
      ) {
        notifyAffectation({
          type: 'anomalie',
          email: affecteFormEmail,
          subject: `T-${result.data.ID}${form.field_4 ? ' — ' + stripHtml(form.field_4).slice(0, 80) : ''}`,
          details: 'Délai de traitement par défaut : 3 jour(s). Voir le détail dans ReportingDCPO.',
        })
      }

      // Notification Teams à l'AUTEUR de l'anomalie (personne à l'origine du
      // fait signalé) — envoyée en fire-and-forget dès qu'un auteur est
      // désigné. Court-circuit si l'auteur est le déclarant lui-même
      // (inutile de se notifier soi-même) ou si aucun auteur n'a été saisi.
      if (
        auteurEmail
        && result.data?.ID
        && auteurEmail.toLowerCase() !== (userEmail ?? '').toLowerCase()
      ) {
        notifyAuteurAnomalie({
          email: auteurEmail,
          anomalieLabel: `T-${result.data.ID}${form.field_4 ? ' — ' + stripHtml(form.field_4).slice(0, 80) : ''}`,
          declareParName: userName,
        })
      }

      resetForm()
      setShowForm(false)
      await fetchItems()
    } catch (err) { console.error('Erreur creation:', err) }
    finally { setSubmitting(false) }
  }

  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   *
   * Structure :
   *   1. Header avec titre + bouton "Nouvelle anomalie"
   *   2. Stats cards (Total, Ouvert, En cours, Resolu, Clos, Montant)
   *   3. Barre de filtres
   *   4. Formulaire de création (toggle via showForm)
   *   5. Tableau principal avec colonnes dépliables
   *   6. Modales (Détail, Affecter, Suivi du ticket, Clôture résolution, Statut)
   * ════════════════════════════════════════════════════════════════════════ */
  return (
    <>
      {/* ─── HEADER : titre + boutons export + "Nouvelle anomalie" ───── */}
      <div className="content-header">
        <h2>Liste des anomalies</h2>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <ExportButtons
            filename="anomalies_en_cours"
            pdfTitle="Liste des anomalies en cours"
            getHeaders={() => [
              'Numéro', 'Titre', 'Statut', 'Criticité', 'Classification',
              'Agence', 'Réseau', 'Domaine activité', 'Cause', 'Date déclaration',
              'Date ouverture', 'Date régularisation', 'Date clôture',
              'Déclarant', 'Auteur', 'Personne affectée', 'Montant',
            ]}
            getRows={() => filteredItems.map(it => [
              it.ID ? `T-${it.ID}` : '',
              it.Title ?? '',
              it.field_10 ?? '',
              it.criticiteAnomalie ?? '',
              it.field_5 ?? '',
              findAgenceLabel(agences, it.field_6),
              findReseauLabel(reseaux, it.field_7),
              it.domaineActivite ?? '',
              it.field_4 ? stripHtml(it.field_4) : '',
              formatDateForExport(it.field_0),
              formatDateForExport(it.dateOuvertureTicket),
              formatDateForExport(it.field_9),
              formatDateForExport(it.date_cloture_ticket),
              it.declarant_anormalie?.DisplayName ?? '',
              it.auteur_anormalie?.DisplayName ?? '',
              it.personneAffecter?.DisplayName ?? '',
              it.field_8 ?? '',
            ])}
            disabled={loading}
          />
          <button
            type="button"
            className={`btn-cta btn-cta-primary btn-cta-add ${showForm ? 'is-active' : ''}`}
            onClick={() => setShowForm(!showForm)}
            aria-expanded={showForm}
          >
            {showForm ? '× Annuler' : '+ Nouvelle anomalie'}
          </button>
        </div>
      </div>

      {/* ─── STATS CARDS : compteurs basés sur filteredItems ─────────── */}
      {/* Toutes les stats reflètent les filtres APPLIQUÉS (cf. filteredItems) */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{filteredItems.length}</span>
          <span className="stat-label">Total</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'Ouvert').length}</span>
          <span className="stat-label">Ouvert</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'En cours').length}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'Resolu').length}</span>
          <span className="stat-label">Resolu</span>
        </div>
        <div className="stat-card clos">
          <span className="stat-value">{filteredItems.filter(i => i.field_10 === 'Clos').length}</span>
          <span className="stat-label">Clos</span>
        </div>
        <div className="stat-card montant">
          {/* Format compact (millions au-delà d'1 M) pour éviter que les très
              gros montants débordent de la carte. */}
          <span className="stat-value">{formatMontantCompact(filteredItems.reduce((s, i) => s + (i.field_8 ?? 0), 0))}</span>
          <span className="stat-label">Montant total</span>
        </div>
      </div>

      {/* ─── BARRE DE FILTRES ──────────────────────────────────────── */}
      {/* Les inputs sont liés à filter* (saisie libre).
          Le bouton Rechercher fait le snapshot vers applied* + refetch. */}
      <div className="filters-bar">
        <div className="filter-field" style={{ flex: '1 1 220px' }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="N°, titre, description, classification..."
            value={filterSearch}
            onChange={e => setFilterSearch(e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') handleSearch() }}
            aria-label="Recherche libre"
          />
        </div>
        <div className="filter-field">
          <label>Date du</label>
          <input type="date" value={filterDateFrom} onChange={e => setFilterDateFrom(e.target.value)} />
        </div>
        <div className="filter-field">
          <label>Date au</label>
          <input type="date" value={filterDateTo} onChange={e => setFilterDateTo(e.target.value)} />
        </div>
        <div className="filter-field">
          <label>Agence</label>
          <select value={filterAgence} onChange={e => setFilterAgence(e.target.value)}>
            <option value="">Toutes</option>
            {agences.map(a => (
              <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Reseau</label>
          <select value={filterReseau} onChange={e => setFilterReseau(e.target.value)}>
            <option value="">Tous</option>
            {reseaux.map(r => (
              <option key={r.ID} value={String(r.ID)}>{r.field_1 ?? r.Title}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Classification</label>
          <select value={filterClassification} onChange={e => setFilterClassification(e.target.value)}>
            <option value="">Toutes</option>
            {CLASSIFICATION_OPTIONS.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Criticite</label>
          <select value={filterCriticite} onChange={e => setFilterCriticite(e.target.value)}>
            <option value="">Toutes</option>
            {CRITICITE_OPTIONS.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Domaine d'activité</label>
          <select value={filterDomaineActivite} onChange={e => setFilterDomaineActivite(e.target.value)}>
            <option value="">Tous</option>
            {DOMAINE_ACTIVITE_OPTIONS.map(d => (
              <option key={d} value={d}>{d}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Statut</label>
          {/* Clos / Résolu retirés volontairement : ils sont consultables sur
              la page "Fiche récapitulatif de l'anomalie". Seuls les tickets
              en cours de traitement apparaissent ici. */}
          <select value={filterStatut} onChange={e => setFilterStatut(e.target.value)}>
            <option value="">Tous (en cours)</option>
            <option value="Ouvert">Ouvert</option>
            <option value="En cours">En cours</option>
          </select>
        </div>
        <div className="filter-field">
          <label>Agent (auteur)</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filterAgent}
            onChange={e => setFilterAgent(e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Personne affectée</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filterAffecte}
            onChange={e => setFilterAffecte(e.target.value)}
          />
        </div>
        <div className="filter-field">
          <label>Déclarant</label>
          <input
            type="text"
            placeholder="Nom ou email..."
            value={filterDeclarant}
            onChange={e => setFilterDeclarant(e.target.value)}
          />
        </div>
        <button className="btn-search-filters" onClick={handleSearch} disabled={loading}>
          {loading ? 'Recherche...' : 'Rechercher'}
        </button>
        <button className="btn-reset-filters" onClick={handleResetFilters}>Reinitialiser</button>
      </div>

      {/* ─── FORMULAIRE DE CRÉATION (toggle via showForm) ──────────── */}
      {/* Organisé en 5 fieldsets thématiques pour clarté UX :
          Identité / Caractérisation / Localisation / Impact / Pièce jointe */}
      {showForm && (
        <form className="create-form modern-form" onSubmit={handleSubmit}>
          <header className="modern-form-header">
            <div>
              <h2>Nouvelle anomalie</h2>
              <p className="modern-form-subtitle">
                Renseignez les informations ci-dessous. Un ticket de suivi sera automatiquement ouvert.
              </p>
            </div>
            <button type="button" className="btn-cta btn-cta-ghost" onClick={() => setShowForm(false)}>
              Annuler
            </button>
          </header>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">👤</span> Identité</legend>
            <div className="form-grid">
              <div className="form-field">
                <label>Déclarant</label>
                <input type="text" readOnly value={userName ?? ''} />
                {userEmail && <span className="selected-email">{userEmail}</span>}
              </div>

              <div className="form-field">
                <label htmlFor="anom-auteur">Auteur</label>
                <div className="autocomplete-wrapper">
                  <input
                    id="anom-auteur"
                    type="text"
                    value={auteurSearch}
                    placeholder="Rechercher un auteur..."
                    onChange={e => searchAuteur(e.target.value)}
                    onBlur={() => setTimeout(() => setShowAuteurDropdown(false), 200)}
                  />
                  {auteurEmail && <span className="selected-email">{auteurEmail}</span>}
                  {showAuteurDropdown && auteurResults.length > 0 && (
                    <ul className="autocomplete-dropdown">
                      {auteurResults.map(u => (
                        <li key={u.Id} onClick={() => selectAuteur(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>

              {/* Champ "Personne affectée" : disponible pour TOUS les rôles.
                  Pré-rempli avec le déclarant lui-même (défaut métier :
                  le contrôleur qui déclare prend en charge le traitement).
                  L'utilisateur peut modifier pour affecter à un collègue,
                  ou vider le champ pour laisser l'anomalie non affectée. */}
              <div className="form-field">
                <label htmlFor="anom-affecte">Personne affectée</label>
                <div className="autocomplete-wrapper">
                  <input
                    id="anom-affecte"
                    type="text"
                    value={affecteFormSearch}
                    placeholder="Rechercher une personne..."
                    onChange={e => searchAffecteForm(e.target.value)}
                    onBlur={() => setTimeout(() => setShowAffecteFormDropdown(false), 200)}
                  />
                  {affecteFormEmail && <span className="selected-email">{affecteFormEmail}</span>}
                  {showAffecteFormDropdown && affecteFormResults.length > 0 && (
                    <ul className="autocomplete-dropdown">
                      {affecteFormResults.map(u => (
                        <li key={u.Id} onClick={() => selectAffecteForm(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
                <small className="field-hint">
                  Par défaut : vous-même. Vous pouvez modifier pour affecter à un(e) collègue.
                </small>
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">📋</span> Caractérisation</legend>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="anom-date">Date de l'anomalie</label>
                <input
                  id="anom-date"
                  type="date"
                  value={form.field_0}
                  max={new Date().toISOString().split('T')[0]}
                  onChange={e => handleChange('field_0', e.target.value)}
                />
              </div>
              <div className="form-field">
                <label htmlFor="anom-classification">Classification</label>
                <select
                  id="anom-classification"
                  value={form.field_5}
                  onChange={e => handleChange('field_5', e.target.value)}
                >
                  <option value="">— Choisir —</option>
                  {CLASSIFICATION_OPTIONS.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="anom-criticite">Criticité</label>
                <select
                  id="anom-criticite"
                  value={form.criticiteAnomalie}
                  onChange={e => handleChange('criticiteAnomalie', e.target.value)}
                >
                  <option value="">— Choisir —</option>
                  {CRITICITE_OPTIONS.map(opt => (
                    <option key={opt} value={opt}>{opt}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="anom-domaine">Domaine d'activité</label>
                <select
                  id="anom-domaine"
                  value={form.domaineActivite}
                  onChange={e => handleChange('domaineActivite', e.target.value)}
                >
                  <option value="">— Choisir —</option>
                  {DOMAINE_ACTIVITE_OPTIONS.map(d => (
                    <option key={d} value={d}>{d}</option>
                  ))}
                </select>
              </div>
              <div className="form-field form-field-full">
                <label htmlFor="anom-cause">Cause / description</label>
                <textarea
                  id="anom-cause"
                  rows={3}
                  value={form.field_4}
                  onChange={e => handleChange('field_4', e.target.value)}
                  placeholder="Décrivez la nature de l'anomalie observée..."
                />
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">📍</span> Localisation</legend>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="anom-agence">Agence</label>
                <select
                  id="anom-agence"
                  value={form.field_6}
                  onChange={e => {
                    const agenceId = e.target.value
                    handleChange('field_6', agenceId)
                    const agence = agences.find(a => String(a.ID) === agenceId)
                    handleChange('field_7', agence?.field_1 ? String(agence.field_1) : '')
                  }}
                >
                  <option value="">— Choisir une agence —</option>
                  {agences.map(a => (
                    <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="anom-reseau">Réseau</label>
                <input
                  id="anom-reseau"
                  type="text"
                  readOnly
                  value={reseaux.find(r => String(r.ID) === form.field_7)?.field_1 ?? ''}
                  placeholder="Sélectionner une agence d'abord"
                />
              </div>
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">💰</span> Impact &amp; suivi</legend>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="anom-montant">Montant</label>
                <input
                  id="anom-montant"
                  type="number"
                  min={0}
                  value={form.field_8}
                  onChange={e => handleChange('field_8', Number(e.target.value))}
                />
              </div>
              <div className="form-field">
                <label htmlFor="anom-statut">Statut initial</label>
                <select
                  id="anom-statut"
                  value={form.field_10}
                  onChange={e => handleChange('field_10', e.target.value)}
                >
                  <option value="">— Choisir (Ouvert par défaut) —</option>
                  <option value="Ouvert">Ouvert</option>
                  <option value="En cours">En cours</option>
                  <option value="Resolu">Résolu</option>
                  <option value="Clos">Clos</option>
                </select>
              </div>
              {/* Date de régularisation : volontairement MASQUÉE à la
                  création — elle n'a de sens qu'au moment de la clôture
                  (cf. modale "Clore la résolution" qui la demande alors).
                  Reste modifiable en édition depuis la modale Détail. */}
            </div>
          </fieldset>

          <fieldset className="form-section">
            <legend><span className="form-section-icon" aria-hidden="true">📎</span> Pièce jointe</legend>
            <div className="form-grid">
              <div className="form-field form-field-full">
                <label htmlFor="anom-attachment">Fichier (optionnel)</label>
                <input
                  id="anom-attachment"
                  type="file"
                  onChange={e => setAttachment(e.target.files?.[0] ?? null)}
                />
                {attachment && <span className="selected-email">📎 {attachment.name}</span>}
              </div>
            </div>
          </fieldset>

          <div className="modern-form-actions">
            <button type="button" className="btn-cta btn-cta-ghost" onClick={() => setShowForm(false)} disabled={submitting}>
              Annuler
            </button>
            <button className="btn-cta btn-cta-primary" type="submit" disabled={submitting}>
              {submitting ? 'Enregistrement...' : 'Créer l\'anomalie'}
            </button>
          </div>
        </form>
      )}

      {/* ─── TABLEAU PRINCIPAL ─────────────────────────────────────── */}
      {/* 3 états : loading / vide / tableau dépliable
          Colonnes toujours visibles : Numéro, Statut, Date, Criticité, Actions
          Colonnes dépliables (toggle "+") : Déclarant, Auteur, Affectée,
            Cause, Classification, Agence, Réseau, Montant, Date régularisation
          Sticky : Numéro à gauche, Actions à droite */}
      {loading ? (
        <p className="loading-text">Chargement des anomalies...</p>
      ) : filteredItems.length === 0 ? (
        <p className="loading-text">Aucune anomalie trouvée.</p>
      ) : (
        <div className="table-wrapper" ref={tableRef}>
          <table className={`data-table anomalies-table ${expandedColumns ? 'is-expanded' : 'is-compact'}`}>
            <thead>
              <tr>
                <th className="col-ticket">Numéro</th>
                <th className="col-status">Statut</th>
                <th className="col-date">Date</th>
                <th className="col-criticite">Criticité</th>
                <th>Délai</th>
                <th className="col-toggle" aria-label="Déplier / replier les colonnes">
                  <button
                    type="button"
                    className="col-toggle-btn"
                    onClick={() => setExpandedColumns(v => !v)}
                    aria-pressed={expandedColumns}
                    aria-label={expandedColumns ? 'Replier les colonnes' : 'Déplier les colonnes'}
                    title={expandedColumns ? 'Replier les colonnes' : 'Déplier les colonnes'}
                  >
                    {expandedColumns ? '−' : '+'}
                  </button>
                </th>
                {expandedColumns && (
                  <>
                    <th className="col-collapsible">Déclarant</th>
                    <th className="col-collapsible">Auteur</th>
                    <th className="col-collapsible">Personne affectée</th>
                    <th className="col-collapsible">Cause</th>
                    <th className="col-collapsible">Classification</th>
                    <th className="col-collapsible">Agence</th>
                    <th className="col-collapsible">Réseau</th>
                    <th className="col-collapsible">Montant</th>
                    <th className="col-collapsible">Date régularisation</th>
                  </>
                )}
                <th className="col-actions">Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedItems.map(item => (
                <tr
                  key={item.ID}
                  className="row-clickable"
                  onClick={() => setDetailItem(item)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setDetailItem(item) } }}
                >
                  <td className="col-ticket">
                    {/* Le numéro de ticket utilise l'ID SharePoint pour être
                        IDENTIQUE à celui affiché sur la page "Fiche
                        récapitulatif de l'anomalie" (cf. buildConsolidatedBulletin
                        qui produit `T-${ticket.ID}`). On évite ainsi qu'un
                        même ticket soit "T-3" ici et "T-46" ailleurs. */}
                    <span
                      className={`ticket-badge ${statusBadgeClass(item.field_10)}`}
                      aria-label={`Ticket T-${item.ID ?? '—'}, statut ${item.field_10 ?? 'inconnu'}`}
                    >
                      <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(item.field_10)}</span>
                      <span className="ticket-badge-num">T-{item.ID ?? '—'}</span>
                    </span>
                  </td>
                  <td className="col-status">
                    <span className={`status-chip ${statusBadgeClass(item.field_10)}`}>
                      {item.field_10 ?? '—'}
                    </span>
                  </td>
                  <td className="col-date">{formatDateOnlyFR(item.field_0, '—')}</td>
                  <td className="col-criticite">
                    {item.criticiteAnomalie
                      ? <span className={`criticite-chip crit-${item.criticiteAnomalie.toLowerCase()}`}>{item.criticiteAnomalie}</span>
                      : '—'}
                  </td>
                  {/* Délai d'affectation : pastille colorée selon que le
                      contrôleur est dans les temps (vert), à échéance (orange)
                      ou en retard (rouge). Cf. helper getDelaiStatus.
                      title= : info-bulle détaillée pour audit rapide. */}
                  <td>
                    {(() => {
                      const s = getDelaiStatus(item.dateAffection, item.delai)
                      if (s.kind === 'na') return <span style={{ color: '#888' }}>—</span>
                      const palette = s.kind === 'restant'
                        ? { bg: '#d1fae5', fg: '#065f46', border: '#6ee7b7' }
                        : s.kind === 'echu'
                          ? { bg: '#fef3c7', fg: '#92400e', border: '#fcd34d' }
                          : { bg: '#fee2e2', fg: '#991b1b', border: '#fca5a5' }
                      const affDate = formatDateOnlyFR(item.dateAffection, '—')
                      return (
                        <span
                          title={`Affecté le ${affDate} — délai ${item.delai ?? '?'} j`}
                          style={{
                            display: 'inline-block',
                            padding: '2px 8px',
                            fontSize: 11,
                            fontWeight: 700,
                            borderRadius: 10,
                            background: palette.bg,
                            color: palette.fg,
                            border: `1px solid ${palette.border}`,
                            whiteSpace: 'nowrap',
                          }}
                        >
                          {s.label}
                        </span>
                      )
                    })()}
                  </td>
                  <td className="col-toggle" aria-hidden="true"></td>
                  {expandedColumns && (
                    <>
                      <td className="col-collapsible">{item.declarant_anormalie?.DisplayName ?? '—'}</td>
                      <td className="col-collapsible">{item.auteur_anormalie?.DisplayName ?? '—'}</td>
                      <td className="col-collapsible">{item.personneAffecter?.DisplayName ?? '—'}</td>
                      <td className="col-collapsible">{item.field_4 ? stripHtml(item.field_4) : '—'}</td>
                      <td className="col-collapsible">{item.field_5 ?? '—'}</td>
                      <td className="col-collapsible">{agences.find(a => String(a.ID) === item.field_6)?.Title ?? item.field_6 ?? '—'}</td>
                      <td className="col-collapsible">{reseaux.find(r => String(r.ID) === item.field_7)?.field_1 ?? item.field_7 ?? '—'}</td>
                      <td className="col-collapsible">{item.field_8?.toLocaleString('fr-FR') ?? '—'}</td>
                      <td className="col-collapsible">{formatDateOnlyFR(item.field_9, '—')}</td>
                    </>
                  )}
                  {/* La cellule Actions stoppe la propagation du clic pour
                      ne pas ouvrir le détail quand on clique sur un bouton. */}
                  <td className="col-actions" onClick={e => e.stopPropagation()}>
                    <div className="actions-col">
                      {/* Bouton "Affecter" : visible pour les managers sur
                          toutes les lignes, et pour les Controleurs uniquement
                          sur les anomalies dont ils sont déclarants (cf.
                          canAffectAnomaly). */}
                      {canAffectAnomaly(item) && (
                        <button type="button" className="btn-cta btn-cta-affect" onClick={() => setAffectItemId(item.ID ?? null)}>
                          Affecter
                        </button>
                      )}
                      <button type="button" className="btn-cta btn-cta-status" onClick={() => setTicketItem(item)}>
                        Ticket
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          {/* Barre de pagination — taille de page configurable globalement
              via DEFAULT_PAGE_SIZE / PAGE_SIZE_OPTIONS dans Pagination.tsx */}
          <Pagination
            state={pagination}
            total={filteredItems.length}
            itemLabel="anomalies"
          />
        </div>
      )}

      {/* ─── MODALE DÉTAIL ─────────────────────────────────────────── */}
      {/* Deux modes :
            - Lecture seule (par défaut) : <dl><dd> avec valeurs formatées
            - Édition : tous les champs métier deviennent des inputs/selects
          Le toggle se fait via le bouton "Modifier" dans le header
          (visible uniquement si canEditAnomaly(detailItem) === true).

          Permissions (cf. canEditAnomaly) :
            - Chef_Departement / Directeur : édition de TOUTES les anomalies
            - Controleur : édition uniquement si elle lui est affectée
            - Champs personnes (auteur/déclarant/affectée) NE SONT PAS éditables
              ici → la personne affectée se change via le bouton "Affecter" */}
      {detailItem && (
        <ModalOverlay onClose={closeDetailModal}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(560px, 100%)' }}>
            <div className="modal-header" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <h2 style={{ flex: 1 }}>{detailEditMode ? "Modifier l'anomalie" : "Detail de l'anomalie"}</h2>
              {/* Bouton "Modifier" : visible si user a les droits ET pas déjà en édition */}
              {!detailEditMode && canEditAnomaly(detailItem) && (
                <button
                  type="button"
                  className="btn-cta btn-cta-detail"
                  onClick={enterDetailEdit}
                >
                  ✎ Modifier
                </button>
              )}
              <button className="modal-close" onClick={closeDetailModal}>&times;</button>
            </div>
            <div className="modal-body">
              {!detailEditMode ? (
                /* ═══ MODE LECTURE ═══════════════════════════════════════ */
                <dl className="detail-grid">
                  <dt>N° Ticket</dt><dd><strong>{detailItem.ID ? `T-${detailItem.ID}` : '-'}</strong></dd>
                  <dt>Ouverture ticket</dt><dd>{detailItem.dateOuvertureTicket ? new Date(detailItem.dateOuvertureTicket).toLocaleString() : '-'}</dd>
                  <dt>Cloture ticket</dt><dd>{detailItem.date_cloture_ticket ? new Date(detailItem.date_cloture_ticket).toLocaleString() : '-'}</dd>
                  <dt>Declarant</dt><dd>{detailItem.declarant_anormalie?.DisplayName ?? '-'}</dd>
                  <dt>Auteur</dt><dd>{detailItem.auteur_anormalie?.DisplayName ?? '-'}</dd>
                  <dt>Personne affectee</dt><dd>{detailItem.personneAffecter?.DisplayName ?? '-'}</dd>
                  <dt>Date d'affectation</dt>
                  <dd>{formatDateOnlyFR(detailItem.dateAffection, '-')}</dd>
                  <dt>Délai de traitement</dt><dd>{detailItem.delai ? `${detailItem.delai} jour(s)` : '-'}</dd>
                  <dt>Commentaire affectation</dt><dd>{detailItem.commentaireAffectation ?? '-'}</dd>
                  <dt>Date</dt><dd>{formatDateOnlyFR(detailItem.field_0, '-')}</dd>
                  <dt>Cause</dt>
                  {/* white-space: pre-line → respecte les \n insérés par stripHtml
                      pour séparer "Cause immédiate / Cause racine / Actions menées /
                      Observations" sur des lignes distinctes. */}
                  <dd style={{ whiteSpace: 'pre-line' }}>
                    {detailItem.field_4 ? stripHtml(detailItem.field_4) : '-'}
                  </dd>
                  <dt>Classification</dt><dd>{detailItem.field_5 ?? '-'}</dd>
                  <dt>Domaine d'activité</dt><dd>{detailItem.domaineActivite ?? '-'}</dd>
                  <dt>Mode de traitement</dt><dd>{detailItem.typeSanction ?? '-'}</dd>
                  <dt>Agence</dt><dd>{agences.find(a => String(a.ID) === detailItem.field_6)?.Title ?? detailItem.field_6 ?? '-'}</dd>
                  <dt>Reseau</dt><dd>{reseaux.find(r => String(r.ID) === detailItem.field_7)?.field_1 ?? detailItem.field_7 ?? '-'}</dd>
                  <dt>Montant</dt><dd>{detailItem.field_8?.toLocaleString() ?? '-'}</dd>
                  <dt>Date regularisation</dt><dd>{formatDateOnlyFR(detailItem.field_9, '-')}</dd>
                  <dt>Statut</dt><dd>{detailItem.field_10 ?? '-'}</dd>
                  <dt>Criticite</dt><dd>{detailItem.criticiteAnomalie ?? '-'}</dd>
                </dl>
              ) : (
                /* ═══ MODE ÉDITION ═══════════════════════════════════════ */
                /* Champs métier éditables. Les champs personnes (auteur,
                   déclarant, personne affectée) restent en lecture seule —
                   personne affectée se change via le bouton "Affecter" dédié. */
                <div className="detail-edit-form">
                  {/* Infos système (toujours lecture seule) */}
                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label>N° Ticket</label>
                    <input type="text" readOnly value={detailItem.ID ? `T-${detailItem.ID}` : '-'} />
                  </div>
                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label>Personne affectée</label>
                    <input type="text" readOnly value={detailItem.personneAffecter?.DisplayName ?? '-'} />
                    <span style={{ fontSize: 11, color: '#888' }}>
                      Pour changer la personne affectée, utiliser le bouton « Affecter ».
                    </span>
                  </div>

                  {/* Champs métier éditables */}
                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-date">Date de l'anomalie</label>
                    <input
                      id="edit-date"
                      type="date"
                      value={detailForm.field_0}
                      max={new Date().toISOString().split('T')[0]}
                      onChange={e => updateDetailForm('field_0', e.target.value)}
                      disabled={detailSaving}
                    />
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-classification">Classification</label>
                    <select
                      id="edit-classification"
                      value={detailForm.field_5}
                      onChange={e => updateDetailForm('field_5', e.target.value)}
                      disabled={detailSaving}
                    >
                      <option value="">— Choisir —</option>
                      {CLASSIFICATION_OPTIONS.map(c => (
                        <option key={c} value={c}>{c}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-criticite">Criticité</label>
                    <select
                      id="edit-criticite"
                      value={detailForm.criticiteAnomalie}
                      onChange={e => updateDetailForm('criticiteAnomalie', e.target.value)}
                      disabled={detailSaving}
                    >
                      <option value="">— Choisir —</option>
                      {CRITICITE_OPTIONS.map(opt => (
                        <option key={opt} value={opt}>{opt}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-domaine">Domaine d'activité</label>
                    <select
                      id="edit-domaine"
                      value={detailForm.domaineActivite}
                      onChange={e => updateDetailForm('domaineActivite', e.target.value)}
                      disabled={detailSaving}
                    >
                      <option value="">— Choisir —</option>
                      {DOMAINE_ACTIVITE_OPTIONS.map(d => (
                        <option key={d} value={d}>{d}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-cause">Cause / description</label>
                    <textarea
                      id="edit-cause"
                      rows={3}
                      value={detailForm.field_4}
                      onChange={e => updateDetailForm('field_4', e.target.value)}
                      disabled={detailSaving}
                    />
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-agence">Agence</label>
                    <select
                      id="edit-agence"
                      value={detailForm.field_6}
                      onChange={e => {
                        const agenceId = e.target.value
                        updateDetailForm('field_6', agenceId)
                        // Auto-déduction du réseau depuis l'agence (cohérent avec la
                        // logique du formulaire de création)
                        const agence = agences.find(a => String(a.ID) === agenceId)
                        updateDetailForm('field_7', agence?.field_1 ? String(agence.field_1) : '')
                      }}
                      disabled={detailSaving}
                    >
                      <option value="">— Choisir une agence —</option>
                      {agences.map(a => (
                        <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
                      ))}
                    </select>
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-reseau">Réseau</label>
                    <input
                      id="edit-reseau"
                      type="text"
                      readOnly
                      value={reseaux.find(r => String(r.ID) === detailForm.field_7)?.field_1 ?? ''}
                      placeholder="Sélectionner une agence d'abord"
                    />
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-montant">Montant</label>
                    <input
                      id="edit-montant"
                      type="number"
                      min={0}
                      value={detailForm.field_8}
                      onChange={e => updateDetailForm('field_8', Number(e.target.value))}
                      disabled={detailSaving}
                    />
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-dateRegul">Date régularisation</label>
                    <input
                      id="edit-dateRegul"
                      type="date"
                      value={detailForm.field_9}
                      onChange={e => updateDetailForm('field_9', e.target.value)}
                      disabled={detailSaving}
                    />
                  </div>

                  {/* Statut : éditable UNIQUEMENT par les managers.
                      Un contrôleur peut éditer les autres champs de "son"
                      anomalie mais pas le statut workflow — celui-ci suit
                      des transitions métier réservées (cf. modale "Changer
                      le statut" et "Clore la résolution"). */}
                  {userRole !== 'Controleur' && (
                    <div className="form-field" style={{ marginBottom: 8 }}>
                      <label htmlFor="edit-statut">Statut</label>
                      <select
                        id="edit-statut"
                        value={detailForm.field_10}
                        onChange={e => updateDetailForm('field_10', e.target.value)}
                        disabled={detailSaving}
                      >
                        <option value="">— Choisir —</option>
                        <option value="Ouvert">Ouvert</option>
                        <option value="En cours">En cours</option>
                        <option value="Resolu">Résolu</option>
                        <option value="Clos">Clos</option>
                      </select>
                    </div>
                  )}

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-delai">Délai de traitement (jours)</label>
                    <input
                      id="edit-delai"
                      type="number"
                      min={1}
                      step={1}
                      value={detailForm.delai}
                      placeholder="Ex: 7"
                      onChange={e => updateDetailForm('delai', e.target.value)}
                      disabled={detailSaving}
                    />
                  </div>

                  <div className="form-field" style={{ marginBottom: 8 }}>
                    <label htmlFor="edit-commentaire">Commentaire d'affectation</label>
                    <textarea
                      id="edit-commentaire"
                      rows={3}
                      value={detailForm.commentaireAffectation}
                      onChange={e => updateDetailForm('commentaireAffectation', e.target.value)}
                      disabled={detailSaving}
                    />
                  </div>

                  {detailError && (
                    <p style={{ color: 'var(--rdcpo-red)', fontSize: 13, margin: '8px 0' }} role="alert">
                      {detailError}
                    </p>
                  )}

                  <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                    <button
                      type="button"
                      className="btn-cta btn-cta-detail"
                      onClick={cancelDetailEdit}
                      disabled={detailSaving}
                    >
                      Annuler
                    </button>
                    <button
                      type="button"
                      className="btn-cta btn-cta-affect"
                      onClick={saveDetailEdit}
                      disabled={detailSaving}
                    >
                      {detailSaving ? 'Enregistrement...' : 'Enregistrer'}
                    </button>
                  </div>
                </div>
              )}

              <div className="detail-attachments" style={{ marginTop: 16 }}>
                <h3 style={{ marginBottom: 8 }}>Pièces jointes</h3>
                {(() => {
                  const attachments = getTicketAttachments(detailItem)
                  if (attachments.length === 0) {
                    return <p className="loading-text">Aucune pièce jointe.</p>
                  }
                  return (
                    <ul className="detail-attachments-list">
                      {attachments.map((att, i) => (
                        <li key={i} className="detail-attachment-item">
                          {att.isImage ? (
                            <img
                              src={att.url}
                              alt={att.name}
                              style={{ width: 56, height: 56, objectFit: 'cover', borderRadius: 6, border: '1px solid #ddd', flexShrink: 0 }}
                            />
                          ) : (
                            <span aria-hidden="true" style={{ fontSize: 24, width: 56, textAlign: 'center', flexShrink: 0 }}>
                              {getAttachmentIcon(att.iconType)}
                            </span>
                          )}
                          <a href={att.url} target="_blank" rel="noreferrer" style={{ color: '#1d4ed8', textDecoration: 'none', fontWeight: 600, wordBreak: 'break-all' }}>
                            {att.name}
                          </a>
                        </li>
                      ))}
                    </ul>
                  )
                })()}
              </div>
            </div>
          </div>
        </ModalOverlay>
      )}

      {/* ─── MODALE AFFECTATION ────────────────────────────────────── */}
      {/* Workflow en 2 étapes :
          1. Recherche Office 365 → sélection d'un utilisateur (setAffectSelectedUser)
          2. Saisie du délai + commentaire → clic "Valider" (confirmAffect)
          → écrit personneAffecter + delai + commentaireAffectation en SharePoint */}
      {affectItemId !== null && (
        <ModalOverlay onClose={closeAffectModal}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>Affecter l'anomalie</h2>
              <button className="modal-close" onClick={closeAffectModal}>&times;</button>
            </div>
            <div className="modal-body">
              {/* ─── ÉTAPE 1 : Recherche utilisateur (tant qu'aucun sélectionné) */}
              {!affectSelectedUser && (
                <>
                  <label htmlFor="affect-search" style={{ fontSize: 12, fontWeight: 700, color: '#666' }}>
                    Personne à affecter
                  </label>
                  <input
                    id="affect-search"
                    type="text"
                    value={affectSearch}
                    placeholder="Rechercher un utilisateur..."
                    onChange={e => searchAffectUser(e.target.value)}
                    autoFocus
                  />
                  {affectResults.length > 0 && (
                    <ul className="modal-results">
                      {affectResults.map(u => (
                        <li key={u.Id} onClick={() => selectAffectUser(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </>
              )}

              {/* ─── ÉTAPE 2 : Formulaire délai + commentaire (après sélection) */}
              {affectSelectedUser && (
                <div className="affect-form">
                  {/* Récap de la personne sélectionnée + bouton "changer" */}
                  <div className="affect-selected">
                    <div>
                      <strong>{affectSelectedUser.DisplayName}</strong>
                      <div style={{ fontSize: 12, color: '#666' }}>{affectSelectedUser.Mail}</div>
                    </div>
                    <button
                      type="button"
                      className="btn-cta btn-cta-detail"
                      onClick={clearAffectSelection}
                      disabled={affectLoading}
                    >
                      Changer
                    </button>
                  </div>

                  {/* Délai : nombre de JOURS (entier > 0). Stocké comme string
                      en SP (typage de la colonne) mais représente bien un entier. */}
                  <div className="form-field" style={{ marginTop: 12 }}>
                    <label htmlFor="affect-delai">Délai de traitement (jours) *</label>
                    <input
                      id="affect-delai"
                      type="number"
                      min={1}
                      step={1}
                      value={affectDelai}
                      placeholder="Ex: 7"
                      onChange={e => setAffectDelai(e.target.value)}
                      disabled={affectLoading}
                    />
                  </div>

                  {/* Commentaire d'affectation : instructions/contexte pour la personne */}
                  <div className="form-field" style={{ marginTop: 8 }}>
                    <label htmlFor="affect-commentaire">Commentaire d'affectation *</label>
                    <textarea
                      id="affect-commentaire"
                      rows={4}
                      value={affectCommentaire}
                      placeholder="Instructions, contexte, points d'attention..."
                      onChange={e => setAffectCommentaire(e.target.value)}
                      disabled={affectLoading}
                    />
                  </div>

                  {affectError && (
                    <p style={{ color: 'var(--rdcpo-red)', fontSize: 13, margin: '8px 0 0' }} role="alert">
                      {affectError}
                    </p>
                  )}

                  <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                    <button
                      type="button"
                      className="btn-cta btn-cta-detail"
                      onClick={closeAffectModal}
                      disabled={affectLoading}
                    >
                      Annuler
                    </button>
                    <button
                      type="button"
                      className="btn-cta btn-cta-affect"
                      onClick={confirmAffect}
                      disabled={affectLoading || !(parseInt(affectDelai, 10) > 0) || !affectCommentaire.trim()}
                    >
                      {affectLoading ? 'Affectation...' : "Valider l'affectation"}
                    </button>
                  </div>
                </div>
              )}

              {affectLoading && !affectSelectedUser && <p className="loading-text">Affectation en cours...</p>}
            </div>
          </div>
        </ModalOverlay>
      )}

      {/* ─── MODALE SUIVI DU TICKET (ouverte par bouton "Ticket") ───── */}
      {/* Vue synthétique avec 3 actions : détail complet, changer statut,
          clore la résolution (cachée si déjà Resolu/Clos) */}
      {ticketItem !== null && (
        <ModalOverlay onClose={() => setTicketItem(null)}>
          <div className="modal ticket-modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <div className="ticket-modal-title">
                <span className={`ticket-badge ticket-badge-lg ${statusBadgeClass(ticketItem.field_10)}`}>
                  <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(ticketItem.field_10)}</span>
                  <span className="ticket-badge-num">T-{ticketItem.ID}</span>
                </span>
                <h2>Suivi du ticket</h2>
              </div>
              <button className="modal-close" onClick={() => setTicketItem(null)} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body ticket-modal-body">
              <div className="ticket-stat-row">
                <span className={`status-chip ${statusBadgeClass(ticketItem.field_10)}`}>
                  {ticketItem.field_10 ?? '—'}
                </span>
                {ticketItem.criticiteAnomalie && (
                  <span className={`criticite-chip crit-${ticketItem.criticiteAnomalie.toLowerCase()}`}>
                    {ticketItem.criticiteAnomalie}
                  </span>
                )}
                {ticketItem.field_5 && <span className="ticket-tag">{ticketItem.field_5}</span>}
              </div>

              <dl className="detail-grid">
                <dt>Ouverture</dt>
                <dd>{ticketItem.dateOuvertureTicket ? new Date(ticketItem.dateOuvertureTicket).toLocaleString('fr-FR') : '—'}</dd>
                <dt>Date déclaration</dt>
                <dd>{formatDateOnlyFR(ticketItem.field_0, '—')}</dd>
                <dt>Régularisation</dt>
                <dd>{formatDateOnlyFR(ticketItem.field_9, '—')}</dd>
                <dt>Clôture</dt>
                <dd>{ticketItem.date_cloture_ticket ? new Date(ticketItem.date_cloture_ticket).toLocaleString('fr-FR') : '—'}</dd>
                <dt>Déclarant</dt>
                <dd>{ticketItem.declarant_anormalie?.DisplayName ?? '—'}</dd>
                <dt>Auteur</dt>
                <dd>{ticketItem.auteur_anormalie?.DisplayName ?? '—'}</dd>
                <dt>Affecté à</dt>
                <dd>{ticketItem.personneAffecter?.DisplayName ?? '—'}</dd>
                <dt>Délai</dt>
                <dd>{ticketItem.delai ? `${ticketItem.delai} jour(s)` : '—'}</dd>
                <dt>Agence</dt>
                <dd>{agences.find(a => String(a.ID) === ticketItem.field_6)?.Title ?? '—'}</dd>
                <dt>Réseau</dt>
                <dd>{reseaux.find(r => String(r.ID) === ticketItem.field_7)?.field_1 ?? '—'}</dd>
                <dt>Montant</dt>
                <dd>{ticketItem.field_8?.toLocaleString('fr-FR') ?? '—'}</dd>
              </dl>

              {ticketItem.field_4 && (
                <div className="ticket-description">
                  <h3>Description / cause</h3>
                  {/* pre-line : préserve les sauts de section (Cause immédiate,
                      Cause racine, Actions menées, Observations). */}
                  <p style={{ whiteSpace: 'pre-line' }}>{stripHtml(ticketItem.field_4)}</p>
                </div>
              )}

              {/* Commentaire d'affectation : visible si renseigné — donne le
                  contexte / les instructions laissés par le manager à la
                  personne assignée au moment de l'affectation. */}
              {ticketItem.commentaireAffectation && (
                <div className="ticket-description">
                  <h3>Commentaire d'affectation</h3>
                  <p>{ticketItem.commentaireAffectation}</p>
                </div>
              )}

              <div className="ticket-modal-actions">
                <button
                  type="button"
                  className="btn-cta btn-cta-secondary"
                  onClick={() => { setTicketItem(null); setDetailItem(ticketItem) }}
                >
                  Voir le détail complet
                </button>
                {/* "Changer le statut" : RÉSERVÉ AUX MANAGERS.
                    Un contrôleur ne peut pas modifier le statut workflow d'un
                    ticket — il peut seulement le clôturer via le bouton
                    "Clore la résolution" (si l'anomalie lui est affectée,
                    cf. canCloseAnomaly).
                    Cohérent avec canAffect : les opérations qui changent
                    l'état métier d'un ticket sont managériales. */}
                {userRole !== 'Controleur' && (
                  <button
                    type="button"
                    className="btn-cta btn-cta-status"
                    onClick={() => { const t = ticketItem; setTicketItem(null); openStatusModal(t) }}
                  >
                    Changer le statut
                  </button>
                )}
                {ticketItem.field_10 !== 'Clos' && ticketItem.field_10 !== 'Resolu' && canCloseAnomaly(ticketItem) && (
                  <button
                    type="button"
                    className="btn-cta btn-cta-primary btn-cta-resolve"
                    onClick={() => openResolution(ticketItem)}
                  >
                    ✓ Clore la résolution
                  </button>
                )}
              </div>
            </div>
          </div>
        </ModalOverlay>
      )}

      {/* ─── MODALE CLÔTURE DE LA RÉSOLUTION ───────────────────────── */}
      {/* Workflow complet contrôleur : statut, dates, auteur (modifiable),
          causes, actions menées (obligatoire), observations, pièce jointe */}
      {resolutionItem !== null && (
        <ModalOverlay onClose={closeResolution}>
          <form className="modal resolution-modal" onClick={e => e.stopPropagation()} onSubmit={saveResolution}>
            <div className="modal-header">
              <div className="ticket-modal-title">
                <span className={`ticket-badge ticket-badge-lg ${statusBadgeClass(resolutionForm.statut)}`}>
                  <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(resolutionForm.statut)}</span>
                  <span className="ticket-badge-num">T-{resolutionItem.ID}</span>
                </span>
                <h2>Clôture de la résolution</h2>
              </div>
              <button type="button" className="modal-close" onClick={closeResolution} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              <p className="resolution-intro">
                Renseignez les actions menées par le contrôleur pour résoudre cette anomalie.
                Le bulletin sera automatiquement consultable dans le sous-menu « Fiche récapitulatif de l'anomalie » après enregistrement.
              </p>

              <div className="resolution-grid">
                <div className="form-field">
                  <label htmlFor="resolution-statut">Statut final</label>
                  <select
                    id="resolution-statut"
                    value={resolutionForm.statut}
                    onChange={e => updateResolutionForm('statut', e.target.value)}
                  >
                    <option value="Resolu">Résolu</option>
                    <option value="Clos">Clos</option>
                  </select>
                </div>
                <div className="form-field">
                  <label htmlFor="resolution-dateRegul">Date de régularisation</label>
                  <input
                    id="resolution-dateRegul"
                    type="date"
                    value={resolutionForm.dateRegul}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => updateResolutionForm('dateRegul', e.target.value)}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="resolution-dateCloture">Date de clôture du ticket</label>
                  <input
                    id="resolution-dateCloture"
                    type="date"
                    value={resolutionForm.dateCloture}
                    min={resolutionItem.dateOuvertureTicket ? resolutionItem.dateOuvertureTicket.split('T')[0] : undefined}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => updateResolutionForm('dateCloture', e.target.value)}
                  />
                </div>
              </div>

              <div className="form-field">
                <label htmlFor="resolution-auteur">Auteur de l'anomalie</label>
                <div className="autocomplete-wrapper">
                  <input
                    id="resolution-auteur"
                    type="text"
                    value={resolutionAuteurSearch}
                    placeholder="Rechercher un auteur..."
                    onChange={e => searchResolutionAuteur(e.target.value)}
                    onBlur={() => setTimeout(() => setShowResolutionAuteurDropdown(false), 200)}
                  />
                  {resolutionAuteurEmail && (
                    <span className="selected-email">{resolutionAuteurEmail}</span>
                  )}
                  {showResolutionAuteurDropdown && resolutionAuteurResults.length > 0 && (
                    <ul className="autocomplete-dropdown">
                      {resolutionAuteurResults.map(u => (
                        <li key={u.Id} onClick={() => selectResolutionAuteur(u)}>
                          <strong>{u.DisplayName}</strong><span>{u.Mail}</span>
                        </li>
                      ))}
                    </ul>
                  )}
                </div>
              </div>

              <div className="form-field">
                <label htmlFor="resolution-causeImmediate">Cause immédiate</label>
                <textarea
                  id="resolution-causeImmediate"
                  rows={2}
                  value={resolutionForm.causeImmediate}
                  onChange={e => updateResolutionForm('causeImmediate', e.target.value)}
                  placeholder="Quelle est la cause directe de l'anomalie ?"
                />
              </div>

              <div className="form-field">
                <label htmlFor="resolution-causeRacine">Cause racine</label>
                <textarea
                  id="resolution-causeRacine"
                  rows={2}
                  value={resolutionForm.causeRacine}
                  onChange={e => updateResolutionForm('causeRacine', e.target.value)}
                  placeholder="Pourquoi cette cause s'est-elle produite (analyse de fond) ?"
                />
              </div>

              <div className="form-field">
                <label htmlFor="resolution-actions">
                  Actions menées <span className="required-mark">*</span>
                  <small className="field-hint" style={{ marginLeft: 8 }}>une action par ligne</small>
                </label>
                <textarea
                  id="resolution-actions"
                  rows={5}
                  value={resolutionForm.actionsMenees}
                  onChange={e => updateResolutionForm('actionsMenees', e.target.value)}
                  placeholder={'Ex :\n- Vérification des écritures\n- Régularisation comptable\n- Notification du déclarant'}
                  required
                />
              </div>

              <div className="form-field">
                <label htmlFor="resolution-observations">Observations</label>
                <textarea
                  id="resolution-observations"
                  rows={2}
                  value={resolutionForm.observations}
                  onChange={e => updateResolutionForm('observations', e.target.value)}
                  placeholder="Remarques additionnelles, recommandations..."
                />
              </div>

              {/* Mode de traitement : saisi UNIQUEMENT à la clôture
                  (pas à la création de l'anomalie). Référentiel mixte
                  TYPE_SANCTION_OPTIONS (relance, demande, sanction…).
                  Champ optionnel — on peut clôturer sans mode de
                  traitement particulier dans les cas simples. */}
              <div className="form-field">
                <label htmlFor="resolution-typeSanction">
                  Mode de traitement
                  <small className="field-hint" style={{ marginLeft: 8 }}>optionnel</small>
                </label>
                <select
                  id="resolution-typeSanction"
                  value={resolutionForm.typeSanction}
                  onChange={e => updateResolutionForm('typeSanction', e.target.value)}
                >
                  <option value="">— Aucun —</option>
                  {TYPE_SANCTION_OPTIONS.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>

              <div className="form-field">
                <label htmlFor="resolution-attachment">
                  Pièce jointe (preuve de résolution)
                  <small className="field-hint" style={{ marginLeft: 8 }}>optionnel</small>
                </label>
                <input
                  id="resolution-attachment"
                  type="file"
                  onChange={e => setResolutionAttachment(e.target.files?.[0] ?? null)}
                />
                {resolutionAttachment && (
                  <span className="selected-email">📎 {resolutionAttachment.name}</span>
                )}
              </div>

              {resolutionError && <p className="status-error" role="alert">{resolutionError}</p>}

              <div className="resolution-actions">
                <button type="button" className="btn-cta btn-cta-secondary" onClick={closeResolution} disabled={resolutionSaving}>
                  Annuler
                </button>
                <button type="submit" className="btn-cta btn-cta-primary btn-cta-resolve" disabled={resolutionSaving}>
                  {resolutionSaving ? 'Enregistrement...' : `Enregistrer (${resolutionForm.statut})`}
                </button>
              </div>
            </div>
          </form>
        </ModalOverlay>
      )}

      {/* ─── MODALE CHANGEMENT DE STATUT (workflow rapide) ─────────── */}
      {/* Permet de basculer le statut sans formulaire complet */}
      {statusItem !== null && (
        <ModalOverlay onClose={closeStatusModal}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(460px, 100%)' }}>
            <div className="modal-header">
              <h2>Changer le statut — Ticket T-{statusItem.ID}</h2>
              <button className="modal-close" onClick={closeStatusModal}>&times;</button>
            </div>
            <div className="modal-body">
              <div className="form-field">
                <label>Statut actuel</label>
                <input type="text" readOnly value={statusItem.field_10 ?? '-'} />
              </div>
              <div className="form-field">
                <label>Nouveau statut</label>
                <select
                  value={statusValue}
                  onChange={e => setStatusValue(e.target.value)}
                  autoFocus
                >
                  <option value="">-- Choisir --</option>
                  <option value="Ouvert">Ouvert</option>
                  <option value="En cours">En cours</option>
                  {/* Resolu / Clos : réservé aux managers ET au contrôleur
                      affecté. Les autres contrôleurs ne voient même pas
                      l'option (cf. canCloseAnomaly). */}
                  {canCloseAnomaly(statusItem) && (
                    <>
                      <option value="Resolu">Résolu</option>
                      <option value="Clos">Clos</option>
                    </>
                  )}
                </select>
              </div>
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <div className="form-field">
                  <label>Date de régularisation</label>
                  <input
                    type="date"
                    value={statusDateRegul}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => setStatusDateRegul(e.target.value)}
                  />
                </div>
              )}
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <div className="form-field">
                  <label>Date de clôture du ticket</label>
                  <input
                    type="date"
                    value={statusDateCloture}
                    min={statusItem?.dateOuvertureTicket ? statusItem.dateOuvertureTicket.split('T')[0] : undefined}
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => setStatusDateCloture(e.target.value)}
                  />
                </div>
              )}
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <p className="status-hint">
                  Le bulletin d'anomalie sera automatiquement disponible dans le sous-menu « Fiche récapitulatif de l'anomalie » après enregistrement.
                </p>
              )}
              {statusError && <p className="status-error" role="alert">{statusError}</p>}
              <div style={{ display: 'flex', gap: 8, marginTop: 16, justifyContent: 'flex-end' }}>
                <button type="button" className="btn-reset-filters" onClick={closeStatusModal} disabled={statusSaving}>
                  Annuler
                </button>
                <button type="button" className="btn-submit" style={{ width: 'auto', marginTop: 0 }} onClick={saveStatus} disabled={statusSaving || !statusValue}>
                  {statusSaving ? 'Enregistrement...' : 'Enregistrer'}
                </button>
              </div>
            </div>
          </div>
        </ModalOverlay>
      )}
    </>
  )
}
