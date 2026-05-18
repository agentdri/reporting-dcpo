/**
 * ============================================================================
 * MODULE 1 — ANOMALIES (Liste, création, gestion ticket)
 * ============================================================================
 *
 * Page centrale de l'application : affiche la liste des anomalies, permet la
 * création, l'affectation, le changement de statut, et la clôture complète
 * via le formulaire de résolution.
 *
 * Fonctionnalités clés :
 *   - Liste filtrable par dates / agence / réseau / classification / criticité /
 *     agent (auteur) / personne affectée
 *   - Filtres serveur (OData) ET client (champs personne) — cf. fetchItems
 *   - Tableau dépliable (colonnes intermédiaires masquables via "+/-")
 *   - Cellules sticky (Numéro à gauche, Actions à droite)
 *   - Badges statut Material Design 3
 *   - Modale création avec sections (Identité / Caractérisation / Localisation /
 *     Impact / Pièce jointe)
 *   - Modale détail (lecture seule)
 *   - Modale affectation (recherche utilisateur Office 365)
 *   - Modale changement de statut rapide
 *   - Modale "Suivi du ticket" (vue synthétique)
 *   - Modale "Clôture de la résolution" (formulaire contrôleur complet :
 *     statut final, dates, auteur, causes, actions, observations, pièce jointe)
 *
 * Persistance :
 *   - Tous les écritures via DCPO_LISTE_ANORMALIEService (SharePoint)
 *   - Pièces jointes : upload via workflow Power Automate puis URL stockée
 *     dans urlPieceJointe (multi-URLs concaténées par "|")
 * ============================================================================
 */

import { useState, useEffect, useMemo } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import { DCPO_LISTE_AGENCESService } from '../generated/services/DCPO_LISTE_AGENCESService'
import { DCPO_LISTE_RESEAUXService } from '../generated/services/DCPO_LISTE_RESEAUXService'
import { Office365UsersService } from '../generated/services/Office365UsersService'
import type { DCPO_LISTE_ANORMALIERead, DCPO_LISTE_ANORMALIEWrite } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import type { User } from '../generated/models/Office365UsersModel'
import { appendUrl, getTicketAttachments, getAttachmentIcon } from '../lib/ticketAttachments'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'

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
}

/** Niveaux de criticité d'anomalie. Utilisé dans les selects. */
const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']

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

/** Nettoie un texte HTML pour ne garder que le texte (DOMParser robuste). */
function stripHtml(html: string): string {
  const doc = new DOMParser().parseFromString(html, 'text/html')
  return doc.body.textContent?.trim() ?? ''
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
   * Règle : seuls les managers (Chef_Departement / Directeur) peuvent
   * assigner une personne sur une anomalie. Un Controleur peut créer et
   * consulter les anomalies mais l'affectation est réservée à sa hiérarchie.
   *
   * On NIE explicitement le rôle Controleur (plutôt que d'autoriser
   * uniquement les managers) pour rester permissif en cas de rôle inconnu
   * ou non renseigné — l'absence de rôle est un cas de bord rare mais
   * possible, et il ne faut pas bloquer un utilisateur légitime à cause
   * d'un trou de configuration.
   */
  const canAffect = userRole !== 'Controleur'

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
  // Filtres client (champs personne, gérés côté frontend)
  const [filterAgent, setFilterAgent] = useState('')
  const [filterAffecte, setFilterAffecte] = useState('')
  // Snapshot des filtres client au moment du clic "Rechercher"
  const [appliedAgent, setAppliedAgent] = useState('')
  const [appliedAffecte, setAppliedAffecte] = useState('')


  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS — MODALES
   * Chaque modale a son propre groupe d'états. La présence d'un objet
   * non-null dans le state principal contrôle l'affichage de la modale.
   * ────────────────────────────────────────────────────────────────────── */

  /** Modale "Détail" (lecture complète d'une anomalie). null = fermée. */
  const [detailItem, setDetailItem] = useState<DCPO_LISTE_ANORMALIERead | null>(null)

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
  const [affectDelai, setAffectDelai] = useState('')
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
  const [affecteFormSearch, setAffecteFormSearch] = useState('')
  const [affecteFormEmail, setAffecteFormEmail] = useState('')
  const [affecteFormResults, setAffecteFormResults] = useState<User[]>([])
  const [showAffecteFormDropdown, setShowAffecteFormDropdown] = useState(false)


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
    setAffectDelai('')
    setAffectCommentaire('')
    setAffectError(null)
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
    setStatusSaving(true)
    setStatusError(null)
    try {
      const payload: Record<string, unknown> = { field_10: statusValue }
      const today = new Date().toISOString().split('T')[0]
      const closesTicket = statusValue === 'Resolu' || statusValue === 'Clos'
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
    if (resolutionForm.actionsMenees.trim()) {
      const items = resolutionForm.actionsMenees
        .split(/\r?\n/)
        .map(s => s.trim())
        .filter(Boolean)
      if (items.length > 1) {
        blocks.push(`<p><strong>Actions menées :</strong></p><ul>${items.map(i => `<li>${escapeHtml(i)}</li>`).join('')}</ul>`)
      } else if (items.length === 1) {
        blocks.push(`<p><strong>Actions menées :</strong> ${escapeHtml(items[0])}</p>`)
      }
    }
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

      // Upload de la pièce jointe (preuve de résolution)
      if (resolutionAttachment) {
        try {
          const attachmentUrl = await uploadAttachment(String(resolutionItem.ID), resolutionAttachment)
          if (attachmentUrl) {
            // Concaténer aux URLs existantes (le champ stocke plusieurs URLs séparées par " | ")
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

      const result = await DCPO_LISTE_ANORMALIEService.getAll(options)
      if (result.data) setItems(result.data)
    } catch (err) { console.error('Erreur chargement anomalies', err) }
    finally { setLoading(false) }
  }

  useEffect(() => {
    const loadLists = async () => {
      try {
        const [agencesRes, reseauxRes] = await Promise.all([
          DCPO_LISTE_AGENCESService.getAll(),
          DCPO_LISTE_RESEAUXService.getAll(),
        ])
        if (agencesRes.data) setAgences(agencesRes.data)
        if (reseauxRes.data) setReseaux(reseauxRes.data)
      } catch (err) { console.error('Erreur chargement agences/reseaux', err) }
    }
    loadLists()
    fetchItems()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

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
    setFilterAgent('')
    setFilterAffecte('')
    setAppliedAgent('')
    setAppliedAffecte('')
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
    if (!agentTerm && !affecteTerm) return items
    return items.filter(it => {
      if (agentTerm) {
        const haystack = `${it.auteur_anormalie?.DisplayName ?? ''} ${it.auteur_anormalie?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(agentTerm)) return false
      }
      if (affecteTerm) {
        const haystack = `${it.personneAffecter?.DisplayName ?? ''} ${it.personneAffecter?.Email ?? ''}`.toLowerCase()
        if (!haystack.includes(affecteTerm)) return false
      }
      return true
    })
  }, [items, appliedAgent, appliedAffecte])

  /* ──────────────────────────────────────────────────────────────────────
   * PAGINATION
   *
   * Le tableau n'affiche que la page courante ; les stats cards restent
   * basées sur filteredItems (vue d'ensemble du filtre, pas de la page).
   * resetKey : remet à la page 1 quand les filtres ou les items changent.
   * ────────────────────────────────────────────────────────────────────── */
  const pagination = usePagination({
    total: filteredItems.length,
    resetKey: `${items.length}|${appliedAgent}|${appliedAffecte}`,
  })
  const pagedItems = useMemo(
    () => filteredItems.slice(pagination.start, pagination.end),
    [filteredItems, pagination.start, pagination.end],
  )


  /* ──────────────────────────────────────────────────────────────────────
   * HANDLERS — FORMULAIRE DE CRÉATION
   * ────────────────────────────────────────────────────────────────────── */

  /** Met à jour un champ du formulaire (utilisé par les inputs). */
  const handleChange = (field: string, value: string | number) => {
    setForm(prev => ({ ...prev, [field]: value }))
  }

  /** Reset complet du formulaire après création réussie ou annulation. */
  const resetForm = () => {
    setForm(EMPTY_FORM)
    setAuteurSearch(''); setAuteurEmail('')
    setAffecteFormSearch(''); setAffecteFormEmail('')
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
      {/* ─── HEADER : titre + bouton "Nouvelle anomalie" ─────────────── */}
      <div className="content-header">
        <h2>Liste des anomalies</h2>
        <button
          type="button"
          className={`btn-cta btn-cta-primary btn-cta-add ${showForm ? 'is-active' : ''}`}
          onClick={() => setShowForm(!showForm)}
          aria-expanded={showForm}
        >
          {showForm ? '× Annuler' : '+ Nouvelle anomalie'}
        </button>
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
          <span className="stat-value">{filteredItems.reduce((s, i) => s + (i.field_8 ?? 0), 0).toLocaleString()}</span>
          <span className="stat-label">Montant total</span>
        </div>
      </div>

      {/* ─── BARRE DE FILTRES ──────────────────────────────────────── */}
      {/* Les inputs sont liés à filter* (saisie libre).
          Le bouton Rechercher fait le snapshot vers applied* + refetch. */}
      <div className="filters-bar">
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
            <option value="Operationnel">Operationnel</option>
            <option value="Fraude">Fraude</option>
            <option value="Commercial">Commercial</option>
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

              {/* Champ "Personne affectée" : réservé aux managers (cf. canAffect).
                  Un Controleur crée une anomalie sans pouvoir l'assigner — c'est
                  ensuite le manager qui réalisera l'affectation depuis le tableau
                  via le bouton "Affecter". */}
              {canAffect && (
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
                </div>
              )}
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
                  <option value="Operationnel">Opérationnel</option>
                  <option value="Fraude">Fraude</option>
                  <option value="Commercial">Commercial</option>
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
              <div className="form-field">
                <label htmlFor="anom-dateRegul">Date régularisation</label>
                <input
                  id="anom-dateRegul"
                  type="date"
                  value={form.field_9}
                  onChange={e => handleChange('field_9', e.target.value)}
                />
              </div>
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
        <div className="table-wrapper">
          <table className={`data-table anomalies-table ${expandedColumns ? 'is-expanded' : 'is-compact'}`}>
            <thead>
              <tr>
                <th className="col-ticket">Numéro</th>
                <th className="col-status">Statut</th>
                <th className="col-date">Date</th>
                <th className="col-criticite">Criticité</th>
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
              {pagedItems.map((item, index) => (
                <tr key={item.ID}>
                  <td className="col-ticket">
                    <span
                      className={`ticket-badge ${statusBadgeClass(item.field_10)}`}
                      aria-label={`Ticket #${pagination.start + index + 1}, statut ${item.field_10 ?? 'inconnu'}`}
                    >
                      <span className="ticket-badge-icon" aria-hidden="true">{statusIcon(item.field_10)}</span>
                      <span className="ticket-badge-num">T-{pagination.start + index + 1}</span>
                    </span>
                  </td>
                  <td className="col-status">
                    <span className={`status-chip ${statusBadgeClass(item.field_10)}`}>
                      {item.field_10 ?? '—'}
                    </span>
                  </td>
                  <td className="col-date">{item.field_0 ? new Date(item.field_0).toLocaleDateString('fr-FR') : '—'}</td>
                  <td className="col-criticite">
                    {item.criticiteAnomalie
                      ? <span className={`criticite-chip crit-${item.criticiteAnomalie.toLowerCase()}`}>{item.criticiteAnomalie}</span>
                      : '—'}
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
                      <td className="col-collapsible">{item.field_9 ? new Date(item.field_9).toLocaleDateString('fr-FR') : '—'}</td>
                    </>
                  )}
                  <td className="col-actions">
                    <div className="actions-col">
                      <button type="button" className="btn-cta btn-cta-detail" onClick={() => setDetailItem(item)}>
                        Détail
                      </button>
                      {/* Bouton "Affecter" : réservé aux managers (cf. canAffect).
                          Un Controleur ne voit pas ce bouton — il consulte mais
                          ne peut pas réassigner l'anomalie. */}
                      {canAffect && (
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

      {/* ─── MODALE DÉTAIL (lecture seule) ─────────────────────────── */}
      {/* Affiche tous les champs de l'anomalie + pièces jointes via getTicketAttachments */}
      {detailItem && (
        <div className="modal-overlay" onClick={() => setDetailItem(null)}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 560 }}>
            <div className="modal-header">
              <h2>Detail de l'anomalie</h2>
              <button className="modal-close" onClick={() => setDetailItem(null)}>&times;</button>
            </div>
            <div className="modal-body">
              <dl className="detail-grid">
                <dt>N° Ticket</dt><dd><strong>{detailItem.ID ? `T-${detailItem.ID}` : '-'}</strong></dd>
                <dt>Ouverture ticket</dt><dd>{detailItem.dateOuvertureTicket ? new Date(detailItem.dateOuvertureTicket).toLocaleString() : '-'}</dd>
                <dt>Cloture ticket</dt><dd>{detailItem.date_cloture_ticket ? new Date(detailItem.date_cloture_ticket).toLocaleString() : '-'}</dd>
                <dt>Declarant</dt><dd>{detailItem.declarant_anormalie?.DisplayName ?? '-'}</dd>
                <dt>Auteur</dt><dd>{detailItem.auteur_anormalie?.DisplayName ?? '-'}</dd>
                <dt>Personne affectee</dt><dd>{detailItem.personneAffecter?.DisplayName ?? '-'}</dd>
                <dt>Délai de traitement</dt><dd>{detailItem.delai ? `${detailItem.delai} jour(s)` : '-'}</dd>
                <dt>Commentaire affectation</dt><dd>{detailItem.commentaireAffectation ?? '-'}</dd>
                <dt>Date</dt><dd>{detailItem.field_0 ? new Date(detailItem.field_0).toLocaleDateString() : '-'}</dd>
                <dt>Cause</dt><dd>{detailItem.field_4 ? stripHtml(detailItem.field_4) : '-'}</dd>
                <dt>Classification</dt><dd>{detailItem.field_5 ?? '-'}</dd>
                <dt>Agence</dt><dd>{agences.find(a => String(a.ID) === detailItem.field_6)?.Title ?? detailItem.field_6 ?? '-'}</dd>
                <dt>Reseau</dt><dd>{reseaux.find(r => String(r.ID) === detailItem.field_7)?.field_1 ?? detailItem.field_7 ?? '-'}</dd>
                <dt>Montant</dt><dd>{detailItem.field_8?.toLocaleString() ?? '-'}</dd>
                <dt>Date regularisation</dt><dd>{detailItem.field_9 ? new Date(detailItem.field_9).toLocaleDateString() : '-'}</dd>
                <dt>Statut</dt><dd>{detailItem.field_10 ?? '-'}</dd>
                <dt>Criticite</dt><dd>{detailItem.criticiteAnomalie ?? '-'}</dd>
              </dl>

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
        </div>
      )}

      {/* ─── MODALE AFFECTATION ────────────────────────────────────── */}
      {/* Workflow en 2 étapes :
          1. Recherche Office 365 → sélection d'un utilisateur (setAffectSelectedUser)
          2. Saisie du délai + commentaire → clic "Valider" (confirmAffect)
          → écrit personneAffecter + delai + commentaireAffectation en SharePoint */}
      {affectItemId !== null && (
        <div className="modal-overlay" onClick={closeAffectModal}>
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
                    <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0 0' }} role="alert">
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
        </div>
      )}

      {/* ─── MODALE SUIVI DU TICKET (ouverte par bouton "Ticket") ───── */}
      {/* Vue synthétique avec 3 actions : détail complet, changer statut,
          clore la résolution (cachée si déjà Resolu/Clos) */}
      {ticketItem !== null && (
        <div className="modal-overlay" onClick={() => setTicketItem(null)}>
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
                <dd>{ticketItem.field_0 ? new Date(ticketItem.field_0).toLocaleDateString('fr-FR') : '—'}</dd>
                <dt>Régularisation</dt>
                <dd>{ticketItem.field_9 ? new Date(ticketItem.field_9).toLocaleDateString('fr-FR') : '—'}</dd>
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
                  <p>{stripHtml(ticketItem.field_4)}</p>
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
                <button
                  type="button"
                  className="btn-cta btn-cta-status"
                  onClick={() => { const t = ticketItem; setTicketItem(null); openStatusModal(t) }}
                >
                  Changer le statut
                </button>
                {ticketItem.field_10 !== 'Clos' && ticketItem.field_10 !== 'Resolu' && (
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
        </div>
      )}

      {/* ─── MODALE CLÔTURE DE LA RÉSOLUTION ───────────────────────── */}
      {/* Workflow complet contrôleur : statut, dates, auteur (modifiable),
          causes, actions menées (obligatoire), observations, pièce jointe */}
      {resolutionItem !== null && (
        <div className="modal-overlay" onClick={closeResolution}>
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
                Le bulletin sera automatiquement consultable dans le sous-menu « Bulletins d'anomalies » après enregistrement.
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
        </div>
      )}

      {/* ─── MODALE CHANGEMENT DE STATUT (workflow rapide) ─────────── */}
      {/* Permet de basculer le statut sans formulaire complet */}
      {statusItem !== null && (
        <div className="modal-overlay" onClick={closeStatusModal}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 460 }}>
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
                  <option value="Resolu">Résolu</option>
                  <option value="Clos">Clos</option>
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
                    max={new Date().toISOString().split('T')[0]}
                    onChange={e => setStatusDateCloture(e.target.value)}
                  />
                </div>
              )}
              {(statusValue === 'Resolu' || statusValue === 'Clos') && (
                <p className="status-hint">
                  Le bulletin d'anomalie sera automatiquement disponible dans le sous-menu « Bulletins d'anomalies » après enregistrement.
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
        </div>
      )}
    </>
  )
}
