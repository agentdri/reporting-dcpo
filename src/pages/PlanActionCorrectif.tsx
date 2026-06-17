/**
 * ============================================================================
 * MODULE — PLAN D'ACTION CORRECTIF (PAC)
 * ============================================================================
 *
 * DEUX CONCEPTS DISTINCTS (architecture similaire à PlanControle.tsx) :
 *
 *   1. PAC (entité principale, liste DCPO_LISTE_PLAN_ACTION_CORRECTIF)
 *      = un plan d'action correctif à mettre en œuvre suite à un constat
 *      (intitulé, sources, descriptions du problème, causes immédiate/racine,
 *      actions correctives, directions concernées, échéance, KPI,
 *      responsable de mise en œuvre, statut).
 *
 *   2. ÉVALUATION (entité dépendante, liste DCPO_EVALUATION_PAC)
 *      = un point de suivi périodique sur l'avancement d'un PAC (observations
 *      + pièces jointes éventuelles). Permet de tracer l'évolution avant la
 *      clôture finale (statut Exécutée / Non Exécutée).
 *
 *      Règle métier : un PAC en statut Exécutée OU Non Exécutée est FIGÉ
 *      (cf. canEvaluatePac dans pacService) — plus d'évaluation possible.
 *
 * STRUCTURE UI
 * ------------
 *   - Stats cards : Total / En cours / Exécutées / Non Exécutées
 *   - Barre de filtres : statut, direction, année + bouton "Nouveau PAC"
 *   - Tableau paginé avec actions (Évaluation, Affectation manager-only)
 *   - Modale détail (lecture seule + historique des évaluations)
 *   - Modale création / édition (managers uniquement)
 *   - Modale ÉVALUATION : créer une évaluation pour un PAC
 *   - Modale AFFECTATION : changer le responsable de mise en œuvre
 *
 * PERMISSIONS
 * -----------
 *   - Chef_Departement / Directeur (canManage) : voient tout, créent, éditent,
 *     affectent, évaluent
 *   - Controleur : voit UNIQUEMENT ses PAC (filtre côté client sur
 *     responsableEmail), peut évaluer (si PAC pas figé) mais pas créer ni
 *     éditer ni affecter
 *
 * PERSISTANCE
 * -----------
 *   - PAC               : src/lib/pacService.ts → DCPO_LISTE_PLAN_ACTION_CORRECTIF
 *   - Évaluations       : même service → DCPO_EVALUATION_PAC
 *   - Référentiel des DIRECTIONS : maintenant en SharePoint
 *     (DCPO_LISTE_DIRECTIONS) — cf. directionService. Liste pré-remplie au
 *     démarrage si vide.
 *   - Pièces jointes (PAC + évaluation) : Power Automate via ticketAttachments
 *
 * Source des champs : feuille "PAC DCPO" du fichier Excel
 * Tableau_de_bord_KPI_DCPO_2026.xlsx fourni par l'utilisateur.
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import {
  createPAC,
  updatePAC,
  updatePacResponsable,
  createPacEvaluation,
  listEvaluationsForPac,
  appendPacEvaluationAttachmentUrls,
  isPacEvaluable,
  type PacEvaluation,
  appendPacAttachmentUrls,
  listPACs,
  getPacStatusClass,
  PAC_STATUS_OPTIONS,
  type Pac,
  type PacStatus,
} from '../lib/pacService'
import { formatDateOnlyFR } from '../lib/formatters'
import { notifyAffectation } from '../lib/teamsNotifications'
import {
  listDirections,
  getDirectionLabelFromList,
  type Direction,
} from '../lib/directionService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { UserPicker } from '../components/UserPicker'
import { uploadPacAttachment, uploadPacEvaluationAttachment, getAttachmentIcon, getAttachmentIconType } from '../lib/ticketAttachments'


/**
 * Props passées par Dashboard.tsx.
 *
 * Seul userRole pilote la logique métier (canManage). Les autres props sont
 * conservées pour cohérence avec les autres modules (logging futur,
 * pré-remplissage responsable, etc.).
 */
interface PlanActionCorrectifProps {
  userName?: string
  userEmail?: string
  userRole?: string
}

/** Rôles ayant le droit de créer / éditer un PAC. */
const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

/** État vide pour la barre de filtres (réutilisé au reset). */
const EMPTY_FILTERS = {
  statut: '' as '' | PacStatus,
  direction: '',
  annee: '',
  search: '',
}

type FilterState = typeof EMPTY_FILTERS

/** État vide pour le formulaire de création. */
const EMPTY_FORM = {
  sourcePac: '',
  dateCreation: new Date().toISOString().split('T')[0],
  intitule: '',
  descriptionProbleme: '',
  causeImmediate: '',
  causeRacine: '',
  actionsCorrectives: '',
  directionsConcernees: [] as string[],
  echeance: '',
  kpi: '',
  annee: new Date().getFullYear(),
  responsableName: '',
  responsableEmail: '',
  statut: 'En cours' as PacStatus,
  observations: '',
}

/**
 * Formatte une date YYYY-MM-DD vers locale FR (jj/mm/aaaa).
 *
 * Délègue à formatDateOnlyFR pour gérer le cas des dates "jour calendaire"
 * stockées comme minuit UTC en SP — évite le décalage de fuseau horaire qui
 * faisait basculer 04/06 → 05/06 selon le TZ du navigateur.
 */
function formatDate(d: string | undefined): string {
  return formatDateOnlyFR(d, '—')
}


export default function PlanActionCorrectif({ userEmail, userRole }: PlanActionCorrectifProps) {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  /** Liste des PAC (chargée depuis SharePoint). */
  const [pacs, setPacs] = useState<Pac[]>([])
  /**
   * Référentiel des directions, chargé depuis la liste SharePoint
   * DCPO_LISTE_DIRECTION (cf. directionService).
   *
   * Démarre vide ; le useEffect d'init le remplit dès la résolution de
   * listDirections(). Tant qu'il est vide, le picker de directions affiche
   * un état "Aucune direction configurée" et les chips PAC affichent les
   * sigles bruts (fallback géré par getDirectionLabelFromList).
   */
  const [directions, setDirections] = useState<Direction[]>([])
  const [loading, setLoading] = useState(true)
  /** Filtres en cours de saisie (binding inputs). */
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** Filtres effectivement appliqués (snapshot au clic Rechercher). */
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  /** PAC sélectionné pour la modale de détail (null = fermée). */
  const [selected, setSelected] = useState<Pac | null>(null)
  /** Affichage du formulaire (création ou édition selon `editingId`). */
  const [showForm, setShowForm] = useState(false)
  /**
   * ID du PAC en cours d'édition (null = mode création).
   * Une seule modale sert aux deux modes — le label et l'action submit
   * s'adaptent en fonction de cet ID.
   */
  const [editingId, setEditingId] = useState<string | null>(null)
  /** Snapshot des champs en cours de saisie dans le formulaire. */
  const [form, setForm] = useState(EMPTY_FORM)
  /** Erreur de soumission affichée sous le formulaire. */
  const [formError, setFormError] = useState<string | null>(null)
  /** True pendant la sauvegarde d'un nouveau PAC. */
  const [saving, setSaving] = useState(false)
  /**
   * Pièce jointe optionnelle attachée au PAC à la création.
   * Uploadée via Power Automate (uploadPacAttachment) APRÈS la création de
   * l'item SP, car le workflow a besoin de l'ID du record pour attacher.
   */
  const [attachment, setAttachment] = useState<File | null>(null)

  /* ─── États de la modale AFFECTATION ────────────────────────────────────
   * Permet à un manager de changer rapidement le responsable de mise en œuvre
   * d'un PAC sans rouvrir le formulaire d'édition complet. */
  /** PAC ciblé par la modale d'affectation (null = modale fermée). */
  const [affectTarget, setAffectTarget] = useState<Pac | null>(null)
  /** Personne sélectionnée dans le picker (nom + email). */
  const [affectSelectedName, setAffectSelectedName] = useState('')
  const [affectSelectedEmail, setAffectSelectedEmail] = useState('')
  /** True pendant l'enregistrement de la nouvelle affectation côté SharePoint. */
  const [affectSaving, setAffectSaving] = useState(false)
  /** Erreur affichée dans la modale (validation ou réseau). */
  const [affectError, setAffectError] = useState<string | null>(null)

  /* ─── États de la modale ÉVALUATION PAC ────────────────────────────────
   * Pas de période (différence avec le Plan de Contrôle) : on consigne juste
   * observations. La modale liste également les évaluations passées du PAC. */
  /** PAC ciblé par la modale d'évaluation (null = modale fermée). */
  const [evalTarget, setEvalTarget] = useState<Pac | null>(null)
  /** Évaluations existantes du PAC ciblé. */
  const [evalHistory, setEvalHistory] = useState<PacEvaluation[]>([])
  /** Champ unique du formulaire (observations). */
  const [evalObservations, setEvalObservations] = useState('')
  /** Pièce jointe optionnelle de l'évaluation PAC.
   *  Uploadée via Power Automate APRÈS création de l'évaluation (le workflow
   *  a besoin de l'ID de l'item), puis l'URL renvoyée est concaténée dans
   *  le champ urlPieceJointe de l'évaluation. */
  const [evalAttachment, setEvalAttachment] = useState<File | null>(null)
  const [evalSaving, setEvalSaving] = useState(false)
  const [evalError, setEvalError] = useState<string | null>(null)

  /** Permission "créer / éditer un PAC". */
  const canManage = !!userRole && MANAGER_ROLES.includes(userRole)


  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT INITIAL + REFRESH (SharePoint)
   * ════════════════════════════════════════════════════════════════════════ */

  // Chargement initial depuis SharePoint. Le fetch async dans un effet est
  // autorisé (le setState a lieu dans un callback après await).
  // On charge PAC et directions en parallèle — pas de dépendance entre eux,
  // et le démarrage est plus rapide qu'en séquence.
  useEffect(() => {
    let cancelled = false
    listPACs()
      .then(data => { if (!cancelled) setPacs(data) })
      .finally(() => { if (!cancelled) setLoading(false) })
    listDirections()
      .then(data => { if (!cancelled) setDirections(data) })
    return () => { cancelled = true }
  }, [])

  /** Recharge la liste depuis SharePoint après création. */
  const refresh = async () => {
    setLoading(true)
    try {
      setPacs(await listPACs())
    } finally {
      setLoading(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — MODALE AFFECTATION (changer le responsable)
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Ouvre la modale d'affectation pour un PAC.
   *
   * Pré-remplit le picker avec la personne actuellement assignée — l'utilisateur
   * peut cliquer "Changer" pour la remplacer (évite la confusion sur qui est
   * affecté actuellement).
   */
  const openAffectation = (pac: Pac) => {
    setAffectTarget(pac)
    setAffectSelectedName(pac.responsable)
    setAffectSelectedEmail(pac.responsableEmail)
    setAffectError(null)
  }

  /** Ferme la modale d'affectation et reset ses états. */
  const closeAffectation = () => {
    setAffectTarget(null)
    setAffectSelectedName('')
    setAffectSelectedEmail('')
    setAffectError(null)
    setAffectSaving(false)
  }

  /**
   * Enregistre la nouvelle affectation côté SharePoint, puis rafraîchit la
   * liste pour que le tableau reflète immédiatement le changement.
   *
   * Garde-fou : on refuse une affectation sans email valide (le champ Person
   * SP nécessite un Claims valide).
   */
  const submitAffectation = async () => {
    if (!affectTarget?.id) return
    setAffectError(null)
    if (!affectSelectedEmail.trim()) {
      setAffectError("Sélectionnez d'abord une personne.")
      return
    }
    setAffectSaving(true)
    try {
      const updated = await updatePacResponsable(affectTarget.id, affectSelectedEmail.trim())
      if (!updated) {
        setAffectError("Échec de l'affectation. Réessayer.")
        return
      }

      // Notification Teams au nouveau responsable de mise en œuvre.
      notifyAffectation({
        type: 'pac',
        email: affectSelectedEmail.trim(),
        subject: affectTarget.intitule || `PAC #${affectTarget.id}`,
        details: `Source : ${affectTarget.sourcePac || '—'}. Échéance : ${affectTarget.echeance || '—'}. Statut : ${affectTarget.statut}.`,
      })

      await refresh()
      closeAffectation()
    } catch (err) {
      console.error('submitAffectation error', err)
      setAffectError(err instanceof Error ? err.message : "Échec de l'affectation.")
    } finally {
      setAffectSaving(false)
    }
  }

  /* ─── Handlers de la modale ÉVALUATION PAC ───────────────────────────── */

  /**
   * Ouvre la modale d'évaluation pour un PAC :
   *   1. Charge l'historique des évaluations existantes (filtre serveur)
   *   2. Reset le formulaire
   *
   * L'historique est chargé en arrière-plan ; pendant ce temps, evalHistory
   * reste vide et l'utilisateur peut déjà commencer à saisir.
   */
  const openEvaluation = async (pac: Pac) => {
    setEvalTarget(pac)
    setEvalObservations('')
    setEvalAttachment(null)
    setEvalError(null)
    setEvalHistory([])
    try {
      const history = await listEvaluationsForPac(pac.id)
      setEvalHistory(history)
    } catch (err) {
      console.error('openEvaluation: échec chargement historique', err)
    }
  }

  /** Ferme la modale d'évaluation et reset tous ses états. */
  const closeEvaluation = () => {
    setEvalTarget(null)
    setEvalObservations('')
    setEvalAttachment(null)
    setEvalHistory([])
    setEvalError(null)
    setEvalSaving(false)
  }

  /**
   * Crée une évaluation pour le PAC ciblé.
   *
   * Garde-fou supplémentaire : on revérifie isPacEvaluable au moment du
   * submit pour empêcher un cas de race où le statut aurait basculé pendant
   * que la modale est ouverte.
   */
  const submitEvaluation = async () => {
    if (!evalTarget?.id) return
    setEvalError(null)

    if (!isPacEvaluable(evalTarget.statut)) {
      setEvalError("Ce PAC n'est plus évaluable (statut Exécutée ou Non Exécutée).")
      return
    }
    if (!evalObservations.trim()) {
      setEvalError('Les observations sont obligatoires.')
      return
    }

    setEvalSaving(true)
    try {
      // 1. Créer l'évaluation côté SP — on récupère son ID pour l'étape 2.
      const created = await createPacEvaluation({
        pacId: evalTarget.id,
        observations: evalObservations.trim(),
      })

      // 2. Si une pièce jointe est sélectionnée : upload via Power Automate,
      //    puis concaténation de l'URL retournée dans urlPieceJointe.
      //
      //    Stratégie d'erreur "best-effort" : si l'upload échoue, on garde
      //    l'évaluation créée et on affiche un message pour ne pas perdre
      //    la saisie utilisateur. L'utilisateur peut re-tenter manuellement
      //    ou attacher le fichier plus tard.
      if (evalAttachment) {
        try {
          const uploadedUrl = await uploadPacEvaluationAttachment(created.id, evalAttachment, 'Visite')
          if (uploadedUrl) {
            await appendPacEvaluationAttachmentUrls(created.id, [uploadedUrl])
          }
        } catch (uploadErr) {
          const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr)
          console.error('Échec upload pièce jointe évaluation PAC', uploadErr)
          setEvalError(`L'évaluation a été enregistrée mais la pièce jointe a échoué : ${detail}`)
          // Recharger l'historique tout de même pour montrer la nouvelle ligne
          const history = await listEvaluationsForPac(evalTarget.id)
          setEvalHistory(history)
          setEvalObservations('')
          setEvalAttachment(null)
          return  // garde la modale ouverte
        }
      }

      // Recharger l'historique pour que la nouvelle évaluation apparaisse en haut.
      const history = await listEvaluationsForPac(evalTarget.id)
      setEvalHistory(history)
      // Reset du formulaire pour permettre une saisie successive
      setEvalObservations('')
      setEvalAttachment(null)
    } catch (err) {
      console.error('submitEvaluation PAC error', err)
      setEvalError(err instanceof Error ? err.message : "Échec de l'enregistrement de l'évaluation.")
    } finally {
      setEvalSaving(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * PAC après application des filtres (côté client, pas d'OData ici).
   *
   * Restriction de visibilité par rôle :
   *   - Controleur → ne voit QUE les PAC dont il est le responsable de mise
   *     en œuvre (responsableEmail = userEmail, case-insensitive)
   *   - Manager (Chef_Departement / Directeur) → voit tout
   *
   * Le filtre rôle est appliqué AVANT les filtres de recherche pour que
   * les compteurs/stats reflètent uniquement ce que l'utilisateur a réellement
   * le droit de voir.
   */
  const filtered = useMemo(() => {
    const myEmail = userEmail?.toLowerCase()
    const restrictToMine = userRole === 'Controleur'
    return pacs.filter(p => {
      // ─── Garde de visibilité par rôle ─────────────────────────────────
      if (restrictToMine) {
        if (!myEmail || p.responsableEmail.toLowerCase() !== myEmail) return false
      }
      // ─── Filtres de la barre de recherche ────────────────────────────
      if (appliedFilters.statut && p.statut !== appliedFilters.statut) return false
      if (appliedFilters.direction && !p.directionsConcernees.includes(appliedFilters.direction)) return false
      if (appliedFilters.annee && String(p.annee) !== appliedFilters.annee) return false
      if (appliedFilters.search) {
        const needle = appliedFilters.search.toLowerCase()
        const haystack = [
          p.intitule,
          p.sourcePac,
          p.descriptionProbleme,
          p.actionsCorrectives,
          p.responsable,
        ].join(' ').toLowerCase()
        if (!haystack.includes(needle)) return false
      }
      return true
    })
  }, [pacs, appliedFilters, userRole, userEmail])

  /** Stats globales pour les cards (recalculées sur la liste filtrée). */
  const stats = useMemo(() => ({
    total: filtered.length,
    enCours: filtered.filter(p => p.statut === 'En cours').length,
    executees: filtered.filter(p => p.statut === 'Exécutée').length,
    nonExecutees: filtered.filter(p => p.statut === 'Non Exécutée').length,
  }), [filtered])

  /**
   * Pagination — porte sur la liste filtrée. resetKey = signature des
   * filtres pour ramener automatiquement à la page 1 quand l'utilisateur
   * change un critère.
   */
  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const pagedPacs = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — FILTRES
   * ════════════════════════════════════════════════════════════════════════ */

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyFilters = () => setAppliedFilters(filters)
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS — FORMULAIRE DE CRÉATION
   * ════════════════════════════════════════════════════════════════════════ */

  /** Helper générique pour patcher un champ du formulaire. */
  const updateForm = <K extends keyof typeof form>(key: K, value: (typeof form)[K]) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  /**
   * Ajoute une direction à la sélection (idempotent : ignore si déjà présent).
   * Appelé depuis le select "Ajouter une direction" du formulaire.
   */
  const addDirection = (code: string) => {
    if (!code) return
    setForm(prev => (
      prev.directionsConcernees.includes(code)
        ? prev
        : { ...prev, directionsConcernees: [...prev.directionsConcernees, code] }
    ))
  }

  /** Retire une direction de la sélection (depuis le × sur le chip). */
  const removeDirection = (code: string) => {
    setForm(prev => ({
      ...prev,
      directionsConcernees: prev.directionsConcernees.filter(c => c !== code),
    }))
  }

  /** Ouvre le formulaire en mode CRÉATION (champs vides). */
  const openForm = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
  }

  /**
   * Ouvre le formulaire en mode ÉDITION (champs pré-remplis depuis un PAC).
   * Le picker UserPicker prend en charge un nom/email pré-rempli via le
   * composant UserPickerEdit (props `initialName` / `initialEmail`).
   */
  const openEdit = (pac: Pac) => {
    setEditingId(pac.id)
    setForm({
      sourcePac: pac.sourcePac,
      dateCreation: pac.dateCreation,
      intitule: pac.intitule,
      descriptionProbleme: pac.descriptionProbleme,
      causeImmediate: pac.causeImmediate,
      causeRacine: pac.causeRacine,
      actionsCorrectives: pac.actionsCorrectives,
      directionsConcernees: pac.directionsConcernees,
      echeance: pac.echeance,
      kpi: pac.kpi,
      annee: pac.annee,
      responsableName: pac.responsable,
      responsableEmail: pac.responsableEmail,
      statut: pac.statut,
      observations: pac.observations ?? '',
    })
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
    // On ferme la modale détail pour que la modale édition soit lisible.
    setSelected(null)
  }

  const closeForm = () => {
    setShowForm(false)
    setEditingId(null)
    setAttachment(null)
    setFormError(null)
  }

  /**
   * Crée OU met à jour le PAC en SharePoint, puis enchaîne (si applicable)
   * l'upload Power Automate de la pièce jointe + persistance de son URL
   * dans le champ `urlPiecesJointes` via `appendPacAttachmentUrls`.
   *
   * Comportement en cas d'échec d'upload :
   *   - Le PAC est déjà créé/modifié en SP → pas de rollback
   *   - On affiche un avertissement et on garde la modale ouverte
   */
  const submitForm = async () => {
    setFormError(null)
    setSaving(true)
    try {
      const input = {
        sourcePac: form.sourcePac.trim(),
        dateCreation: form.dateCreation,
        intitule: form.intitule.trim(),
        descriptionProbleme: form.descriptionProbleme.trim(),
        causeImmediate: form.causeImmediate.trim(),
        causeRacine: form.causeRacine.trim(),
        actionsCorrectives: form.actionsCorrectives.trim(),
        directionsConcernees: form.directionsConcernees,
        echeance: form.echeance,
        kpi: form.kpi.trim(),
        annee: Number(form.annee) || new Date().getFullYear(),
        responsableName: form.responsableName,
        responsableEmail: form.responsableEmail,
        statut: form.statut,
        observations: form.observations.trim() || undefined,
      }

      // Mémo de l'ancien responsable AVANT update — pour ne notifier que
      // si le responsable a effectivement CHANGÉ.
      const previousResponsableEmail = editingId
        ? pacs.find(p => p.id === editingId)?.responsableEmail ?? ''
        : ''

      // Création OU édition selon editingId
      const saved = editingId
        ? await updatePAC(editingId, input)
        : await createPAC(input)

      if (!saved?.id) {
        throw new Error(editingId ? 'Échec de la mise à jour.' : 'Échec de la création.')
      }

      // Notification Teams si :
      //   - création avec responsable renseigné
      //   - OU édition qui change l'email du responsable de mise en œuvre
      const responsableChanged =
        !!input.responsableEmail &&
        input.responsableEmail.toLowerCase() !== previousResponsableEmail.toLowerCase()
      if (responsableChanged) {
        notifyAffectation({
          type: 'pac',
          email: input.responsableEmail,
          subject: saved.intitule || `PAC #${saved.id}`,
          details: `Source : ${saved.sourcePac || '—'}. Échéance : ${saved.echeance || '—'}. Statut : ${saved.statut}.`,
        })
      }

      // Upload optionnel de la pièce jointe via le workflow Power Automate dédié.
      // L'URL renvoyée est ensuite concaténée dans le champ urlPiecesJointes
      // (multi-URLs séparées par " | ", idem anomalies).
      if (attachment && saved.id) {
        try {
          const uploadedUrl = await uploadPacAttachment(saved.id, attachment, 'Visite')
          if (uploadedUrl) {
            await appendPacAttachmentUrls(saved.id, [uploadedUrl])
          }
        } catch (uploadErr) {
          const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr)
          console.error('Échec upload pièce jointe PAC', uploadErr)
          setFormError(`Le PAC a été enregistré mais la pièce jointe a échoué : ${detail}`)
          await refresh()
          return  // on garde la modale ouverte
        }
      }

      await refresh()
      closeForm()
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Échec de l\'enregistrement du PAC.')
    } finally {
      setSaving(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   * ════════════════════════════════════════════════════════════════════════ */

  return (
    <>
      {/* ─── Header ────────────────────────────────────────────────── */}
      <div className="content-header">
        <h2>Plan d'Action Correctif</h2>
        {canManage && (
          <button className="btn-add" type="button" onClick={openForm}>
            + Nouveau PAC
          </button>
        )}
      </div>

      {/* ─── Stats cards ───────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Total PAC</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.enCours}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.executees}</span>
          <span className="stat-label">Exécutées</span>
        </div>
        <div className="stat-card ouvert">
          <span className="stat-value">{stats.nonExecutees}</span>
          <span className="stat-label">Non Exécutées</span>
        </div>
      </div>

      {/* ─── Barre de filtres ──────────────────────────────────────── */}
      <div className="filters-bar">
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Intitulé, description, responsable..."
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <div className="filter-field">
          <label>Statut</label>
          <select value={filters.statut} onChange={e => updateFilter('statut', e.target.value as FilterState['statut'])}>
            <option value="">Tous</option>
            {PAC_STATUS_OPTIONS.map(s => (
              <option key={s} value={s}>{s}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Direction</label>
          <select value={filters.direction} onChange={e => updateFilter('direction', e.target.value)}>
            <option value="">Toutes</option>
            {directions.map(d => (
              <option key={d.sigle} value={d.sigle}>{d.sigle} — {d.libelle}</option>
            ))}
          </select>
        </div>
        <div className="filter-field" style={{ maxWidth: 110 }}>
          <label>Année</label>
          <input
            type="number"
            placeholder="2026"
            value={filters.annee}
            onChange={e => updateFilter('annee', e.target.value)}
          />
        </div>
        <div className="filter-actions">
          <button type="button" className="btn-search-filters" onClick={applyFilters}>
            Rechercher
          </button>
          <button type="button" className="btn-reset-filters" onClick={resetFilters}>
            Réinitialiser
          </button>
        </div>
      </div>

      {/* ─── Tableau des PAC ───────────────────────────────────────── */}
      {loading ? (
        <p className="loading-text" style={{ marginTop: 16 }}>Chargement des PAC...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {pacs.length === 0
            ? 'Aucun PAC enregistré pour le moment.'
            : 'Aucun PAC ne correspond aux critères.'}
        </p>
      ) : (
        <div className="table-wrapper" style={{ marginTop: 16 }}>
          <table className="data-table anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Source</th>
                <th>Date</th>
                <th>Intitulé</th>
                <th>Directions</th>
                <th>Échéance</th>
                <th>Responsable</th>
                <th>Statut</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedPacs.map((p, i) => (
                <tr
                  key={p.id}
                  className="row-clickable"
                  onClick={() => setSelected(p)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setSelected(p) } }}
                >
                  <td>{pagination.start + i + 1}</td>
                  <td>{p.sourcePac || '—'}</td>
                  <td>{formatDate(p.dateCreation)}</td>
                  <td style={{ maxWidth: 280 }}>{p.intitule}</td>
                  <td>
                    {p.directionsConcernees.map(c => (
                      <span key={c} className="ticket-tag" style={{ marginRight: 4 }}>{c}</span>
                    ))}
                  </td>
                  <td>{formatDate(p.echeance)}</td>
                  <td>{p.responsable || '—'}</td>
                  <td>
                    <span className={`manager-pill ${getPacStatusClass(p.statut)}`}>{p.statut}</span>
                  </td>
                  {/* Stop propagation : éviter d'ouvrir le détail en cliquant
                      sur un bouton d'action de la cellule. */}
                  <td onClick={e => e.stopPropagation()}>
                    <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                      {/* Bouton "Évaluation" : visible uniquement si le PAC
                          est encore évaluable (statut "En cours"). Dès qu'il
                          passe à "Exécutée" ou "Non Exécutée", il est figé
                          côté évaluation — l'historique reste consultable
                          mais aucune nouvelle évaluation ne peut être ajoutée. */}
                      {isPacEvaluable(p.statut) && (
                        <button type="button" className="btn-cta btn-cta-affect" onClick={() => openEvaluation(p)}>
                          Évaluation
                        </button>
                      )}
                      {/* Bouton "Affectation" : réservé aux managers (canManage).
                          Un Controleur consulte les PAC mais ne peut pas
                          réassigner le responsable. */}
                      {canManage && (
                        <button type="button" className="btn-cta btn-cta-affect" onClick={() => openAffectation(p)}>
                          Affectation
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
          <Pagination
            state={pagination}
            total={filtered.length}
            itemLabel="PAC"
          />
        </div>
      )}

      {/* ─── Modale de détail (lecture seule) ──────────────────────── */}
      {selected && (
        <PacDetailModal
          pac={selected}
          directions={directions}
          onClose={() => setSelected(null)}
          onEdit={canManage ? () => openEdit(selected) : undefined}
        />
      )}

      {/* ─── Modale AFFECTATION ────────────────────────────────────── */}
      {/* Workflow rapide pour changer le responsable de mise en œuvre d'un PAC
          sans rouvrir le formulaire d'édition complet. Réservé aux managers
          (le bouton qui ouvre cette modale n'est rendu que si canManage).
          Le picker est pré-rempli avec la personne déjà affectée. */}
      {affectTarget && (
        <div className="modal-overlay" onClick={closeAffectation}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(520px, 100%)' }}>
            <div className="modal-header">
              <h2>Affecter un responsable</h2>
              <button className="modal-close" onClick={closeAffectation} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              {/* Rappel du PAC ciblé (lecture seule) */}
              <dl className="detail-grid" style={{ marginBottom: 12 }}>
                <dt>PAC</dt><dd><strong>{affectTarget.intitule}</strong></dd>
                <dt>Source</dt><dd>{affectTarget.sourcePac || '—'}</dd>
                <dt>Responsable actuel</dt>
                <dd>
                  {affectTarget.responsable
                    ? <>{affectTarget.responsable}<span style={{ color: '#888', fontSize: 12 }}> ({affectTarget.responsableEmail})</span></>
                    : '—'}
                </dd>
              </dl>

              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-affect-user">Nouveau responsable *</label>
                <UserPicker
                  id="pac-affect-user"
                  selectedName={affectSelectedName}
                  selectedEmail={affectSelectedEmail}
                  onSelect={(name, email) => {
                    setAffectSelectedName(name)
                    setAffectSelectedEmail(email)
                    setAffectError(null)
                  }}
                  onClear={() => {
                    setAffectSelectedName('')
                    setAffectSelectedEmail('')
                  }}
                  disabled={affectSaving}
                  placeholder="Rechercher une personne Office 365..."
                />
              </div>

              {affectError && (
                <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
                  {affectError}
                </p>
              )}

              <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button
                  type="button"
                  className="btn-cta btn-cta-detail"
                  onClick={closeAffectation}
                  disabled={affectSaving}
                >
                  Annuler
                </button>
                <button
                  type="button"
                  className="btn-cta btn-cta-affect"
                  onClick={submitAffectation}
                  disabled={affectSaving || !affectSelectedEmail.trim()}
                >
                  {affectSaving ? 'Affectation...' : "Valider l'affectation"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ─── Modale ÉVALUATION PAC ─────────────────────────────────── */}
      {/* Permet de consigner une évaluation pour un PAC en cours.
            - Pas de période (différence avec Plan de Contrôle) : un PAC
              est ponctuel, on évalue son avancement à un moment T.
            - Le bouton qui ouvre cette modale n'est rendu que si
              isPacEvaluable(p.statut), mais on revérifie au submit pour
              éviter les races (le statut peut changer pendant l'ouverture).
            - L'historique des évaluations existantes est listé en bas. */}
      {evalTarget && (
        <div className="modal-overlay" onClick={closeEvaluation}>
          <div
            className="modal"
            onClick={e => e.stopPropagation()}
            style={{ width: 'min(640px, 100%)' }}
            role="dialog"
            aria-modal="true"
          >
            <div className="modal-header">
              <h2>Évaluation du PAC</h2>
              <button className="modal-close" onClick={closeEvaluation} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              {/* Rappel du PAC ciblé (lecture seule) */}
              <dl className="detail-grid" style={{ marginBottom: 12 }}>
                <dt>Intitulé</dt><dd><strong>{evalTarget.intitule}</strong></dd>
                <dt>Source</dt><dd>{evalTarget.sourcePac || '—'}</dd>
                <dt>Responsable</dt><dd>{evalTarget.responsable || '—'}</dd>
                <dt>Statut</dt>
                <dd>
                  <span className={`manager-pill ${getPacStatusClass(evalTarget.statut)}`}>
                    {evalTarget.statut}
                  </span>
                </dd>
              </dl>

              <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>Nouvelle évaluation</h3>

              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-eval-obs">Observations *</label>
                <textarea
                  id="pac-eval-obs"
                  rows={4}
                  value={evalObservations}
                  onChange={e => setEvalObservations(e.target.value)}
                  placeholder="Avancement, blocages, jalons franchis..."
                  disabled={evalSaving}
                />
              </div>

              {/* Pièce jointe optionnelle : uploadée après création de
                  l'évaluation (workflow Power Automate). En cas d'échec
                  d'upload, l'évaluation reste créée — message d'avertissement
                  affiché à l'utilisateur. */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-eval-attachment">Pièce jointe (optionnel)</label>
                <input
                  id="pac-eval-attachment"
                  type="file"
                  onChange={e => setEvalAttachment(e.target.files?.[0] ?? null)}
                  disabled={evalSaving}
                />
                {evalAttachment && (
                  <span className="selected-email" style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                    {getAttachmentIcon(getAttachmentIconType(evalAttachment.name))} {evalAttachment.name}
                  </span>
                )}
              </div>

              {evalError && (
                <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
                  {evalError}
                </p>
              )}

              <div
                className="modal-actions"
                style={{ marginTop: 8, display: 'flex', gap: 8, justifyContent: 'flex-end' }}
              >
                <button
                  type="button"
                  className="btn-cta btn-cta-detail"
                  onClick={closeEvaluation}
                  disabled={evalSaving}
                >
                  Fermer
                </button>
                <button
                  type="button"
                  className="btn-cta btn-cta-affect"
                  onClick={submitEvaluation}
                  disabled={evalSaving || !evalObservations.trim()}
                >
                  {evalSaving ? 'Enregistrement...' : "Enregistrer l'évaluation"}
                </button>
              </div>

              {/* Historique des évaluations existantes du PAC */}
              <div style={{ marginTop: 20 }}>
                <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>
                  Historique des évaluations ({evalHistory.length})
                </h3>
                {evalHistory.length === 0 ? (
                  <p className="loading-text">Aucune évaluation enregistrée pour le moment.</p>
                ) : (
                  <ul style={{ listStyle: 'none', padding: 0, margin: 0, display: 'flex', flexDirection: 'column', gap: 8 }}>
                    {evalHistory.map(ev => (
                      <li
                        key={ev.id}
                        style={{
                          padding: '8px 10px',
                          border: '1px solid #eee',
                          borderLeft: '3px solid #1d4ed8',
                          borderRadius: 6,
                          background: '#fafafa',
                        }}
                      >
                        <div style={{ fontSize: 12, color: '#888' }}>
                          {ev.createdAt && new Date(ev.createdAt).toLocaleDateString('fr-FR')}
                        </div>
                        {ev.observations && (
                          <div style={{ marginTop: 4, fontSize: 13, color: '#444', whiteSpace: 'pre-wrap' }}>
                            {ev.observations}
                          </div>
                        )}
                        {/* Pièces jointes éventuelles de l'évaluation */}
                        {ev.attachments.length > 0 && (
                          <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                            {ev.attachments.map((att, i) => (
                              <a
                                key={i}
                                href={att.url}
                                target="_blank"
                                rel="noreferrer"
                                style={{
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: 4,
                                  fontSize: 12,
                                  color: '#1d4ed8',
                                  textDecoration: 'none',
                                  padding: '2px 6px',
                                  border: '1px solid #bfdbfe',
                                  borderRadius: 4,
                                  background: '#fff',
                                }}
                              >
                                {getAttachmentIcon(getAttachmentIconType(att.name))} {att.name}
                              </a>
                            ))}
                          </div>
                        )}
                      </li>
                    ))}
                  </ul>
                )}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ─── Modale de création ────────────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
            <div className="modal-header">
              <h2>{editingId ? 'Modifier le PAC' : "Nouveau Plan d'Action Correctif"}</h2>
              <button className="modal-close" onClick={closeForm}>&times;</button>
            </div>
            <div className="modal-body">
              {/* Champs d'identification */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-intitule">Intitulé du PAC *</label>
                <input
                  id="pac-intitule"
                  type="text"
                  value={form.intitule}
                  onChange={e => updateForm('intitule', e.target.value)}
                  placeholder="Ex: Frais de convention SSP"
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-source">Source PAC</label>
                <input
                  id="pac-source"
                  type="text"
                  value={form.sourcePac}
                  onChange={e => updateForm('sourcePac', e.target.value)}
                  placeholder="Ex: Contrôle administratifs"
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-date">Date d'ouverture</label>
                <input
                  id="pac-date"
                  type="date"
                  value={form.dateCreation}
                  onChange={e => updateForm('dateCreation', e.target.value)}
                />
              </div>

              {/* Diagnostic du problème */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-desc">Description du problème</label>
                <textarea
                  id="pac-desc"
                  rows={2}
                  value={form.descriptionProbleme}
                  onChange={e => updateForm('descriptionProbleme', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-cause-im">Cause immédiate</label>
                <textarea
                  id="pac-cause-im"
                  rows={2}
                  value={form.causeImmediate}
                  onChange={e => updateForm('causeImmediate', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-cause-rac">Cause racine</label>
                <textarea
                  id="pac-cause-rac"
                  rows={2}
                  value={form.causeRacine}
                  onChange={e => updateForm('causeRacine', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-actions">Actions correctives à mener</label>
                <textarea
                  id="pac-actions"
                  rows={2}
                  value={form.actionsCorrectives}
                  onChange={e => updateForm('actionsCorrectives', e.target.value)}
                />
              </div>

              {/* Directions concernées (multi-sélection via dropdown + chips)
                  - Select : on n'affiche que les directions PAS encore sélectionnées
                    (évite l'ajout en double et clarifie ce qui reste à choisir)
                  - Le select revient toujours en value="" après ajout
                    (signal pour React de réafficher le placeholder)
                  - Chips : chaque direction sélectionnée s'affiche avec × pour
                    la retirer (UX standard du multi-select à tags) */}
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="pac-direction-add">Directions concernées *</label>
                {directions.length === 0 ? (
                  <p style={{ color: '#888', fontSize: 12, margin: 0 }}>
                    Aucune direction configurée.
                  </p>
                ) : (
                  <>
                    <select
                      id="pac-direction-add"
                      value=""
                      onChange={e => {
                        addDirection(e.target.value)
                        // Reset visuel : on remet le select sur "Ajouter..." pour
                        // permettre l'ajout successif sans changer manuellement.
                        e.target.value = ''
                      }}
                    >
                      <option value="">
                        {form.directionsConcernees.length === directions.length
                          ? '— Toutes les directions sélectionnées —'
                          : '— Ajouter une direction —'}
                      </option>
                      {directions
                        .filter(d => !form.directionsConcernees.includes(d.sigle))
                        .map(d => (
                          <option key={d.sigle} value={d.sigle}>
                            {d.sigle} — {d.libelle}
                          </option>
                        ))}
                    </select>
                    {/* Chips des directions sélectionnées */}
                    {form.directionsConcernees.length > 0 && (
                      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6, marginTop: 6 }}>
                        {form.directionsConcernees.map(code => {
                          const dir = directions.find(d => d.sigle === code)
                          return (
                            <span
                              key={code}
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: 6,
                                padding: '2px 6px 2px 8px',
                                background: '#e0f2fe',
                                border: '1px solid #bae6fd',
                                borderRadius: 12,
                                fontSize: 12,
                              }}
                            >
                              <strong>{code}</strong>
                              {dir && <span style={{ color: '#444' }}>— {dir.libelle}</span>}
                              <button
                                type="button"
                                onClick={() => removeDirection(code)}
                                aria-label={`Retirer ${code}`}
                                style={{
                                  border: 'none',
                                  background: 'transparent',
                                  cursor: 'pointer',
                                  fontSize: 14,
                                  lineHeight: 1,
                                  color: '#0369a1',
                                  padding: '0 2px',
                                }}
                              >
                                ×
                              </button>
                            </span>
                          )
                        })}
                      </div>
                    )}
                  </>
                )}
              </div>

              {/* Pilotage : échéance, KPI, année, responsable, statut */}
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                <div className="form-field">
                  <label htmlFor="pac-echeance">Échéance</label>
                  <input
                    id="pac-echeance"
                    type="date"
                    value={form.echeance}
                    onChange={e => updateForm('echeance', e.target.value)}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="pac-annee">Année</label>
                  <input
                    id="pac-annee"
                    type="number"
                    value={form.annee}
                    onChange={e => updateForm('annee', Number(e.target.value))}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="pac-resp">Responsable de mise en œuvre</label>
                  {/* Champ Personne SharePoint → sélecteur Office 365 */}
                  <UserPicker
                    id="pac-resp"
                    selectedName={form.responsableName}
                    selectedEmail={form.responsableEmail}
                    onSelect={(name, email) => {
                      updateForm('responsableName', name)
                      updateForm('responsableEmail', email)
                    }}
                    onClear={() => {
                      updateForm('responsableName', '')
                      updateForm('responsableEmail', '')
                    }}
                    disabled={saving}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="pac-statut">Statut initial</label>
                  <select
                    id="pac-statut"
                    value={form.statut}
                    onChange={e => updateForm('statut', e.target.value as PacStatus)}
                  >
                    {PAC_STATUS_OPTIONS.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="pac-kpi">KPI</label>
                <input
                  id="pac-kpi"
                  type="text"
                  value={form.kpi}
                  onChange={e => updateForm('kpi', e.target.value)}
                  placeholder="Ex: 100% des dossiers traités dans les délais"
                />
              </div>
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="pac-obs">Observations initiales</label>
                <textarea
                  id="pac-obs"
                  rows={2}
                  value={form.observations}
                  onChange={e => updateForm('observations', e.target.value)}
                />
              </div>

              {/* Pièce jointe optionnelle — uploadée après création via Power
                  Automate (workflow PAC_ATTACHMENT_API_URL). Attache le fichier
                  à l'item SharePoint nouvellement créé. */}
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="pac-attachment">Pièce jointe (optionnel)</label>
                <input
                  id="pac-attachment"
                  type="file"
                  onChange={e => setAttachment(e.target.files?.[0] ?? null)}
                  disabled={saving}
                />
                {attachment && (
                  <span className="selected-email">📎 {attachment.name}</span>
                )}
              </div>

              {formError && (
                <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
                  {formError}
                </p>
              )}

              <div className="modal-actions" style={{ marginTop: 12, display: 'flex', gap: 8, justifyContent: 'flex-end' }}>
                <button type="button" className="btn-cta btn-cta-detail" onClick={closeForm} disabled={saving}>
                  Annuler
                </button>
                <button type="button" className="btn-cta btn-cta-affect" onClick={submitForm} disabled={saving}>
                  {saving
                    ? (editingId ? 'Enregistrement...' : 'Création...')
                    : (editingId ? 'Enregistrer' : 'Créer le PAC')}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </>
  )
}


/* ══════════════════════════════════════════════════════════════════════════
 * MODALE DÉTAIL (lecture seule)
 *
 * Affiche tous les champs du PAC + ses pièces jointes. Si l'utilisateur a
 * les droits manager (cf. canManage côté parent), le prop `onEdit` est
 * fourni et un bouton "Modifier" est affiché en header pour ouvrir le
 * formulaire en mode édition.
 * ══════════════════════════════════════════════════════════════════════════ */

interface PacDetailModalProps {
  pac: Pac
  /** Référentiel des directions pour résoudre les libellés à partir des sigles. */
  directions: Direction[]
  onClose: () => void
  /** Callback pour basculer en mode édition (undefined = bouton masqué). */
  onEdit?: () => void
}

function PacDetailModal({ pac, directions, onClose, onEdit }: PacDetailModalProps) {
  // Historique des évaluations du PAC : chargé une fois à l'ouverture
  // (filtre serveur OData sur plan_action_correctif_id).
  const [history, setHistory] = useState<PacEvaluation[]>([])
  const [historyLoading, setHistoryLoading] = useState(true)

  // Escape ferme la modale (cohérent avec les autres modales du projet)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  useEffect(() => {
    // Pas de setHistoryLoading(true) ici : l'état initial est déjà `true`.
    // La modale se monte à chaque ouverture donc un useState(true) initial
    // suffit — pas de re-fetch sur même item.
    let cancelled = false
    listEvaluationsForPac(pac.id)
      .then(data => { if (!cancelled) setHistory(data) })
      .finally(() => { if (!cancelled) setHistoryLoading(false) })
    return () => { cancelled = true }
  }, [pac.id])

  const attachments = pac.attachments ?? []

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
        <div className="modal-header" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <h2 style={{ flex: 1 }}>Détail du PAC</h2>
          {/* Bouton Modifier : visible uniquement si onEdit fourni (manager). */}
          {onEdit && (
            <button type="button" className="btn-cta btn-cta-detail" onClick={onEdit}>
              ✎ Modifier
            </button>
          )}
          <button className="modal-close" onClick={onClose}>&times;</button>
        </div>
        <div className="modal-body">
          <dl className="detail-grid">
            <dt>Intitulé</dt><dd><strong>{pac.intitule}</strong></dd>
            <dt>Source</dt><dd>{pac.sourcePac || '—'}</dd>
            <dt>Date d'ouverture</dt><dd>{formatDate(pac.dateCreation)}</dd>
            <dt>Année</dt><dd>{pac.annee}</dd>
            <dt>Statut</dt>
            <dd>
              <span className={`manager-pill ${getPacStatusClass(pac.statut)}`}>{pac.statut}</span>
            </dd>
            <dt>Directions concernées</dt>
            <dd>
              {pac.directionsConcernees.length === 0
                ? '—'
                : pac.directionsConcernees.map(c => (
                    <span key={c} className="ticket-tag" style={{ marginRight: 6 }}>
                      {c} — {getDirectionLabelFromList(directions, c)}
                    </span>
                  ))}
            </dd>
            <dt>Échéance</dt><dd>{formatDate(pac.echeance)}</dd>
            <dt>Responsable</dt><dd>{pac.responsable || '—'}</dd>
            <dt>KPI</dt><dd>{pac.kpi || '—'}</dd>

            <dt>Description du problème</dt><dd>{pac.descriptionProbleme || '—'}</dd>
            <dt>Cause immédiate</dt><dd>{pac.causeImmediate || '—'}</dd>
            <dt>Cause racine</dt><dd>{pac.causeRacine || '—'}</dd>
            <dt>Actions correctives</dt><dd>{pac.actionsCorrectives || '—'}</dd>

            {pac.observations && (<><dt>Observations</dt><dd>{pac.observations}</dd></>)}
          </dl>

          {/* Pièces jointes — icônes selon extension, lien externe vers le fichier */}
          <div className="detail-attachments" style={{ marginTop: 16 }}>
            <h3 style={{ marginBottom: 8 }}>Pièces jointes</h3>
            {attachments.length === 0 ? (
              <p className="loading-text">Aucune pièce jointe.</p>
            ) : (
              <ul className="detail-attachments-list">
                {attachments.map((att, i) => {
                  const iconType = getAttachmentIconType(att.name)
                  return (
                    <li key={i} className="detail-attachment-item">
                      <span aria-hidden="true" style={{ fontSize: 24, width: 56, textAlign: 'center', flexShrink: 0 }}>
                        {getAttachmentIcon(iconType)}
                      </span>
                      <a href={att.url} target="_blank" rel="noreferrer" style={{ color: '#1d4ed8', textDecoration: 'none', fontWeight: 600, wordBreak: 'break-all' }}>
                        {att.name}
                      </a>
                    </li>
                  )
                })}
              </ul>
            )}
          </div>

          {/* Historique des évaluations du PAC — même rendu visuel que dans
              la modale d'évaluation pour la cohérence d'expérience. */}
          <div style={{ marginTop: 20 }}>
            <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>
              Historique des évaluations ({history.length})
            </h3>
            {historyLoading ? (
              <p className="loading-text">Chargement de l'historique...</p>
            ) : history.length === 0 ? (
              <p className="loading-text">Aucune évaluation enregistrée pour ce PAC.</p>
            ) : (
              <ul style={{ listStyle: 'none', padding: 0, margin: 0, display: 'flex', flexDirection: 'column', gap: 8 }}>
                {history.map(ev => (
                  <li
                    key={ev.id}
                    style={{
                      padding: '8px 10px',
                      border: '1px solid #eee',
                      borderLeft: '3px solid #1d4ed8',
                      borderRadius: 6,
                      background: '#fafafa',
                    }}
                  >
                    <div style={{ fontSize: 12, color: '#888' }}>
                      {ev.createdAt && new Date(ev.createdAt).toLocaleDateString('fr-FR')}
                    </div>
                    {ev.observations && (
                      <div style={{ marginTop: 4, fontSize: 13, color: '#444', whiteSpace: 'pre-wrap' }}>
                        {ev.observations}
                      </div>
                    )}
                    {ev.attachments.length > 0 && (
                      <div style={{ marginTop: 6, display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                        {ev.attachments.map((att, i) => (
                          <a
                            key={i}
                            href={att.url}
                            target="_blank"
                            rel="noreferrer"
                            style={{
                              display: 'inline-flex',
                              alignItems: 'center',
                              gap: 4,
                              fontSize: 12,
                              color: '#1d4ed8',
                              textDecoration: 'none',
                              padding: '2px 6px',
                              border: '1px solid #bfdbfe',
                              borderRadius: 4,
                              background: '#fff',
                            }}
                          >
                            {getAttachmentIcon(getAttachmentIconType(att.name))} {att.name}
                          </a>
                        ))}
                      </div>
                    )}
                  </li>
                ))}
              </ul>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}
