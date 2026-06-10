/**
 * ============================================================================
 * MODULE — PLAN DE CONTRÔLE (PLAN ANNUEL DCPO)
 * ============================================================================
 *
 * Permet aux managers de créer et suivre les activités de contrôle prévues
 * sur l'année. Inspiré du fichier Excel "PLAN DE CONTROLE 2025" fourni par
 * l'utilisateur, avec :
 *   - une catégorie (regroupement thématique des contrôles)
 *   - un libellé d'activité
 *   - une fréquence (Quotidienne / Hebdomadaire / Mensuelle / Annuelle)
 *   - un responsable (champ Personne Office 365)
 *   - des objectifs (qualitatif + chiffré/KPI)
 *   - un statut workflow
 *
 * Structure UI :
 *   - Stats cards : Total + compteurs par statut
 *   - Filtres : catégorie, fréquence, responsable, statut, année
 *   - Tableau paginé
 *   - Bouton "Nouveau contrôle" (managers uniquement)
 *   - Modale création
 *   - Modale détail (lecture seule)
 *
 * Persistance : liste SharePoint DCPO_LISTE_PLAN_CONTROLE via
 * src/lib/planControleService.ts.
 * ============================================================================
 */

import { useMemo, useState, useEffect } from 'react'
import {
  createControle,
  updateControle,
  updateControleResponsable,
  appendPlanControleAttachmentUrls,
  listControles,
  getControleStatusClass,
  createEvaluation,
  appendEvaluationAttachmentUrls,
  listEvaluationsForControle,
  listAllEvaluations,
  computeEvaluationProgress,
  getExpectedEvaluationsPerYear,
  getPeriodeInputType,
  formatPeriodeLabel,
  PLAN_CONTROLE_CATEGORIES,
  FREQUENCE_OPTIONS,
  STATUS_OPTIONS,
  type ControleEntry,
  type ControleFrequence,
  type ControleStatus,
  type ControleEvaluation,
} from '../lib/planControleService'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { UserPicker } from '../components/UserPicker'
import { ProgressBar } from '../components/ProgressBar'
import {
  uploadPlanControleAttachment,
  uploadEvaluationAttachment,
  getAttachmentIcon,
  getAttachmentIconType,
} from '../lib/ticketAttachments'


interface PlanControleProps {
  userName?: string
  userEmail?: string
  userRole?: string
}

/** Rôles autorisés à créer / éditer un contrôle. */
const MANAGER_ROLES = ['Chef_Departement', 'Directeur']

/** État vide pour la barre de filtres. */
const EMPTY_FILTERS = {
  categorie: '',
  frequence: '' as '' | ControleFrequence,
  responsable: '',
  statut: '' as '' | ControleStatus,
  annee: '',
  search: '',
}

type FilterState = typeof EMPTY_FILTERS

/** État vide pour le formulaire de création. */
const EMPTY_FORM = {
  libelle: '',
  categorie: PLAN_CONTROLE_CATEGORIES[0],
  objectif: '',
  objectifChiffre: '',
  frequence: 'Mensuelle' as ControleFrequence,
  responsableName: '',
  responsableEmail: '',
  annee: new Date().getFullYear(),
  statut: 'À planifier' as ControleStatus,
}


export default function PlanControle({ userEmail, userRole }: PlanControleProps) {
  /* ════════════════════════════════════════════════════════════════════════
   * ÉTATS
   * ════════════════════════════════════════════════════════════════════════ */

  const [controles, setControles] = useState<ControleEntry[]>([])
  const [loading, setLoading] = useState(true)
  /**
   * Map planControleId → nombre d'évaluations effectuées pour ce contrôle.
   * Sert au calcul du taux d'évolution affiché dans le tableau.
   *
   * On charge TOUTES les évaluations en une seule requête au montage pour
   * éviter N appels (un par contrôle). Le compte est ensuite construit en
   * mémoire — adapté pour des volumes < 5000 évaluations (limite SP par
   * défaut). Si on dépasse, basculer sur une agrégation côté serveur.
   */
  const [evaluationCounts, setEvaluationCounts] = useState<Map<number, number>>(new Map())
  const [filters, setFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [appliedFilters, setAppliedFilters] = useState<FilterState>(EMPTY_FILTERS)
  const [selected, setSelected] = useState<ControleEntry | null>(null)
  const [showForm, setShowForm] = useState(false)
  /**
   * ID du contrôle en cours d'édition (null = mode création).
   * Une seule modale sert aux deux modes.
   */
  const [editingId, setEditingId] = useState<string | null>(null)
  const [form, setForm] = useState(EMPTY_FORM)
  const [formError, setFormError] = useState<string | null>(null)
  const [saving, setSaving] = useState(false)
  /**
   * Pièce jointe optionnelle. À la création comme à la modification : uploadée
   * via Power Automate APRÈS création/update de l'item SP (le workflow a
   * besoin de l'ID), puis l'URL renvoyée est ajoutée au champ urlPieceJointe.
   */
  const [attachment, setAttachment] = useState<File | null>(null)

  /* ─── États de la modale ÉVALUATION ─────────────────────────────────────
   * Une seule modale gère :
   *   - le formulaire de saisie (observations + période)
   *   - l'affichage des évaluations passées du contrôle (historique)
   * Elle s'ouvre via le bouton "Évaluation" dans les actions du tableau. */
  /** Contrôle ciblé par la modale d'évaluation (null = modale fermée). */
  const [evalTarget, setEvalTarget] = useState<ControleEntry | null>(null)
  /** Évaluations existantes du contrôle ciblé (rechargées à l'ouverture). */
  const [evalHistory, setEvalHistory] = useState<ControleEvaluation[]>([])
  /** Champs du formulaire d'évaluation. */
  const [evalForm, setEvalForm] = useState({ periode: '', observations: '' })
  /** True pendant la sauvegarde de l'évaluation. */
  const [evalSaving, setEvalSaving] = useState(false)
  /** Erreur affichée dans la modale (validation ou réseau). */
  const [evalError, setEvalError] = useState<string | null>(null)
  /**
   * Pièce jointe optionnelle d'une évaluation. Uploadée via Power Automate
   * APRÈS création de l'évaluation (le workflow a besoin de l'ID de l'item),
   * puis l'URL renvoyée est ajoutée au champ urlPieceJointe.
   */
  const [evalAttachment, setEvalAttachment] = useState<File | null>(null)

  /* ─── États de la modale AFFECTATION ────────────────────────────────────
   * Permet à un manager de changer rapidement la personne responsable d'un
   * contrôle sans rouvrir le formulaire d'édition complet. */
  /** Contrôle ciblé par la modale d'affectation (null = modale fermée). */
  const [affectTarget, setAffectTarget] = useState<ControleEntry | null>(null)
  /** Personne sélectionnée dans le picker (nom + email). */
  const [affectSelectedName, setAffectSelectedName] = useState('')
  const [affectSelectedEmail, setAffectSelectedEmail] = useState('')
  /** True pendant l'enregistrement de la nouvelle affectation côté SharePoint. */
  const [affectSaving, setAffectSaving] = useState(false)
  /** Erreur affichée dans la modale (validation ou réseau). */
  const [affectError, setAffectError] = useState<string | null>(null)

  /** Permission "créer un contrôle". */
  const canManage = !!userRole && MANAGER_ROLES.includes(userRole)

  /* ════════════════════════════════════════════════════════════════════════
   * CHARGEMENT DEPUIS SHAREPOINT
   * Le fetch async dans un effet est autorisé (setState dans un callback
   * après await, pas dans le corps synchrone de l'effet).
   * ════════════════════════════════════════════════════════════════════════ */

  useEffect(() => {
    let cancelled = false
    // Chargement parallèle : contrôles + toutes les évaluations en une fois.
    // Les évaluations alimentent le compteur affiché dans la colonne
    // "Taux d'évolution" du tableau.
    listControles()
      .then(data => { if (!cancelled) setControles(data) })
      .finally(() => { if (!cancelled) setLoading(false) })
    listAllEvaluations()
      .then(evals => {
        if (cancelled) return
        const map = new Map<number, number>()
        for (const ev of evals) {
          if (!ev.planControleId) continue
          map.set(ev.planControleId, (map.get(ev.planControleId) ?? 0) + 1)
        }
        setEvaluationCounts(map)
      })
    return () => { cancelled = true }
  }, [])


  /* ════════════════════════════════════════════════════════════════════════
   * FILTRAGE + PAGINATION + STATS
   * ════════════════════════════════════════════════════════════════════════ */

  /**
   * Liste filtrée côté client (toutes les colonnes sont en mémoire).
   *
   * Restriction de visibilité par rôle :
   *   - Controleur → ne voit QUE les contrôles dont il est responsable
   *     (responsableEmail = userEmail, case-insensitive)
   *   - Manager (Chef_Departement / Directeur) → voit tout
   *
   * Le filtre rôle est appliqué AVANT les filtres de la barre de recherche
   * pour que les compteurs/stats reflètent uniquement ce que l'utilisateur
   * a réellement le droit de voir.
   */
  const filtered = useMemo(() => {
    const myEmail = userEmail?.toLowerCase()
    const restrictToMine = userRole === 'Controleur'
    return controles.filter(c => {
      // ─── Garde de visibilité par rôle (cf. doc ci-dessus) ─────────────
      if (restrictToMine) {
        if (!myEmail || c.responsableEmail.toLowerCase() !== myEmail) return false
      }
      // ─── Filtres de la barre de recherche ────────────────────────────
      if (appliedFilters.categorie && c.categorie !== appliedFilters.categorie) return false
      if (appliedFilters.frequence && c.frequence !== appliedFilters.frequence) return false
      if (appliedFilters.statut && c.statut !== appliedFilters.statut) return false
      if (appliedFilters.annee && String(c.annee) !== appliedFilters.annee) return false
      if (appliedFilters.responsable) {
        const needle = appliedFilters.responsable.toLowerCase()
        if (!c.responsable.toLowerCase().includes(needle)) return false
      }
      if (appliedFilters.search) {
        const needle = appliedFilters.search.toLowerCase()
        const haystack = `${c.libelle} ${c.objectif} ${c.objectifChiffre}`.toLowerCase()
        if (!haystack.includes(needle)) return false
      }
      return true
    })
  }, [controles, appliedFilters, userRole, userEmail])

  /**
   * Stats globales (sur la liste filtrée) + taux d'évolution global.
   *
   * `tauxGlobal` = ratio agrégé de toutes les évaluations effectuées vs
   * attendues sur les contrôles visibles. Calculé comme :
   *   sum(évaluations faites)  /  sum(évaluations attendues) * 100
   *
   * Pertinent pour donner une vision macro à la direction : un seul
   * pourcentage qui résume "où en est-on collectivement sur le plan de
   * contrôle ?". Capé à 100 % et arrondi à l'entier.
   *
   * Les contrôles sans évaluation attendue (cas pathologique) sont ignorés
   * pour éviter une division par zéro.
   */
  const stats = useMemo(() => {
    let totalRealized = 0
    let totalExpected = 0
    for (const c of filtered) {
      const expected = getExpectedEvaluationsPerYear(c.frequence)
      if (expected <= 0) continue
      totalRealized += evaluationCounts.get(Number(c.id)) ?? 0
      totalExpected += expected
    }
    const tauxGlobal = totalExpected > 0
      ? Math.max(0, Math.min(100, Math.round((totalRealized / totalExpected) * 100)))
      : 0
    return {
      total: filtered.length,
      aPlanifier: filtered.filter(c => c.statut === 'À planifier').length,
      planifies: filtered.filter(c => c.statut === 'Planifié').length,
      enCours: filtered.filter(c => c.statut === 'En cours').length,
      realises: filtered.filter(c => c.statut === 'Réalisé').length,
      tauxGlobal,
      totalRealized,
      totalExpected,
    }
  }, [filtered, evaluationCounts])

  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  const pagedControles = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )


  /* ════════════════════════════════════════════════════════════════════════
   * HANDLERS
   * ════════════════════════════════════════════════════════════════════════ */

  const refresh = async () => {
    setLoading(true)
    try {
      setControles(await listControles())
    } finally {
      setLoading(false)
    }
  }

  const updateFilter = <K extends keyof FilterState>(key: K, value: FilterState[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }
  const applyFilters = () => setAppliedFilters(filters)
  const resetFilters = () => {
    setFilters(EMPTY_FILTERS)
    setAppliedFilters(EMPTY_FILTERS)
  }

  const updateForm = <K extends keyof typeof form>(key: K, value: (typeof form)[K]) => {
    setForm(prev => ({ ...prev, [key]: value }))
  }

  /** Ouvre le formulaire en mode CRÉATION. */
  const openForm = () => {
    setEditingId(null)
    setForm(EMPTY_FORM)
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
  }

  /** Ouvre le formulaire en mode ÉDITION pré-rempli depuis un contrôle. */
  const openEdit = (ctrl: ControleEntry) => {
    setEditingId(ctrl.id)
    setForm({
      libelle: ctrl.libelle,
      categorie: ctrl.categorie,
      objectif: ctrl.objectif,
      objectifChiffre: ctrl.objectifChiffre,
      frequence: ctrl.frequence,
      responsableName: ctrl.responsable,
      responsableEmail: ctrl.responsableEmail,
      annee: ctrl.annee,
      statut: ctrl.statut,
    })
    setAttachment(null)
    setFormError(null)
    setShowForm(true)
    // Ferme la modale détail pour laisser la place à la modale édition.
    setSelected(null)
  }

  const closeForm = () => {
    setShowForm(false)
    setEditingId(null)
    setAttachment(null)
    setFormError(null)
  }

  /* ─── Handlers de la modale ÉVALUATION ────────────────────────────────── */

  /**
   * Ouvre la modale d'évaluation pour un contrôle donné :
   *   1. Charge l'historique des évaluations existantes (filtre serveur sur
   *      plan_controle_id)
   *   2. Reset le formulaire à vide
   *
   * L'historique est chargé en arrière-plan ; pendant ce temps, evalHistory
   * reste vide et l'utilisateur peut déjà commencer à saisir.
   */
  const openEvaluation = async (ctrl: ControleEntry) => {
    setEvalTarget(ctrl)
    setEvalForm({ periode: '', observations: '' })
    setEvalAttachment(null)
    setEvalError(null)
    setEvalHistory([])
    try {
      const history = await listEvaluationsForControle(ctrl.id)
      setEvalHistory(history)
    } catch (err) {
      console.error('openEvaluation: échec chargement historique', err)
    }
  }

  /** Ferme la modale d'évaluation et reset tous ses états. */
  const closeEvaluation = () => {
    setEvalTarget(null)
    setEvalForm({ periode: '', observations: '' })
    setEvalAttachment(null)
    setEvalHistory([])
    setEvalError(null)
    setEvalSaving(false)
  }

  /**
   * Crée une évaluation pour le contrôle ciblé.
   *
   * Règles métier :
   *   - la période est obligatoire (sinon l'évaluation perd son sens temporel)
   *   - les observations sont obligatoires (sinon l'évaluation est vide)
   *
   * Après succès : on recharge l'historique et on vide le formulaire — la
   * modale reste ouverte pour permettre de saisir une autre évaluation sur
   * une autre période sans avoir à rouvrir.
   */
  const submitEvaluation = async () => {
    if (!evalTarget?.id) return
    setEvalError(null)

    if (!evalForm.periode.trim()) {
      setEvalError('La période est obligatoire.')
      return
    }
    if (!evalForm.observations.trim()) {
      setEvalError('Les observations sont obligatoires.')
      return
    }

    setEvalSaving(true)
    try {
      const created = await createEvaluation({
        planControleId: Number(evalTarget.id),
        periode: evalForm.periode.trim(),
        observations: evalForm.observations.trim(),
      })

      // Upload optionnel de la pièce jointe via le workflow Power Automate dédié
      // (même mécanique que pour les anomalies / contrôles / PAC) :
      //   1. uploadEvaluationAttachment encode en base64 + POST au workflow
      //   2. Le workflow attache le fichier à l'item SP et renvoie son URL
      //   3. appendEvaluationAttachmentUrls concatène cette URL dans le champ
      //      urlPieceJointe (multi-URLs séparées par " | ")
      if (evalAttachment && created?.id) {
        try {
          const uploadedUrl = await uploadEvaluationAttachment(created.id, evalAttachment, 'Visite')
          if (uploadedUrl) {
            await appendEvaluationAttachmentUrls(created.id, [uploadedUrl])
          }
        } catch (uploadErr) {
          const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr)
          console.error('Échec upload pièce jointe évaluation', uploadErr)
          // L'évaluation est déjà créée → on recharge l'historique pour la
          // refléter, mais on garde la modale ouverte avec le message d'erreur
          // pour que l'utilisateur sache que la PJ a échoué.
          const history = await listEvaluationsForControle(evalTarget.id)
          setEvalHistory(history)
          setEvalError(`L'évaluation a été créée mais la pièce jointe a échoué : ${detail}`)
          return
        }
      }

      // Recharger l'historique pour que la nouvelle évaluation apparaisse en haut.
      const history = await listEvaluationsForControle(evalTarget.id)
      setEvalHistory(history)
      // Reset du formulaire pour permettre une saisie successive
      setEvalForm({ periode: '', observations: '' })
      setEvalAttachment(null)
    } catch (err) {
      console.error('submitEvaluation error', err)
      setEvalError(err instanceof Error ? err.message : "Échec de l'enregistrement de l'évaluation.")
    } finally {
      setEvalSaving(false)
    }
  }

  /* ─── Handlers de la modale AFFECTATION ──────────────────────────────── */

  /**
   * Ouvre la modale d'affectation pour un contrôle.
   *
   * Initialise le picker avec la personne actuellement assignée (s'il y en a
   * une) — l'utilisateur peut cliquer "Changer" pour la remplacer.
   */
  const openAffectation = (ctrl: ControleEntry) => {
    setAffectTarget(ctrl)
    setAffectSelectedName(ctrl.responsable)
    setAffectSelectedEmail(ctrl.responsableEmail)
    setAffectError(null)
  }

  /** Ferme la modale d'affectation et reset tous ses états. */
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
      const updated = await updateControleResponsable(affectTarget.id, affectSelectedEmail.trim())
      if (!updated) {
        setAffectError("Échec de l'affectation. Réessayer.")
        return
      }
      // Refresh complet de la liste pour que la nouvelle affectation apparaisse
      const refreshed = await listControles()
      setControles(refreshed)
      closeAffectation()
    } catch (err) {
      console.error('submitAffectation error', err)
      setAffectError(err instanceof Error ? err.message : "Échec de l'affectation.")
    } finally {
      setAffectSaving(false)
    }
  }

  /**
   * Crée OU met à jour le contrôle (selon editingId), puis enchaîne l'upload
   * de la pièce jointe et la persistance de son URL dans le champ
   * `urlPieceJointe` via `appendPlanControleAttachmentUrls`.
   *
   * Si la création/update réussit mais l'upload échoue, on garde la modale
   * ouverte avec un message d'erreur — l'item existe déjà en SP, pas de rollback.
   */
  const submitForm = async () => {
    setFormError(null)
    setSaving(true)
    try {
      const input = {
        libelle: form.libelle,
        categorie: form.categorie,
        objectif: form.objectif,
        objectifChiffre: form.objectifChiffre,
        frequence: form.frequence,
        responsableName: form.responsableName,
        responsableEmail: form.responsableEmail,
        annee: Number(form.annee) || new Date().getFullYear(),
        statut: form.statut,
      }

      const saved = editingId
        ? await updateControle(editingId, input)
        : await createControle(input)

      if (!saved?.id) {
        throw new Error(editingId ? 'Échec de la mise à jour.' : 'Échec de la création.')
      }

      // Upload optionnel via workflow Power Automate + persistance de l'URL.
      if (attachment && saved.id) {
        try {
          const uploadedUrl = await uploadPlanControleAttachment(saved.id, attachment, 'Visite')
          if (uploadedUrl) {
            await appendPlanControleAttachmentUrls(saved.id, [uploadedUrl])
          }
        } catch (uploadErr) {
          const detail = uploadErr instanceof Error ? uploadErr.message : String(uploadErr)
          console.error('Échec upload pièce jointe Plan de Contrôle', uploadErr)
          setFormError(`Le contrôle a été enregistré mais la pièce jointe a échoué : ${detail}`)
          await refresh()
          return  // garde la modale ouverte
        }
      }

      await refresh()
      closeForm()
    } catch (err) {
      setFormError(err instanceof Error ? err.message : 'Échec de l\'enregistrement du contrôle.')
    } finally {
      setSaving(false)
    }
  }


  /* ════════════════════════════════════════════════════════════════════════
   * RENDU
   * ════════════════════════════════════════════════════════════════════════ */

  return (
    <>
      <div className="content-header">
        <h2>Plan de Contrôle</h2>
        {canManage && (
          <button className="btn-add" type="button" onClick={openForm}>
            + Nouveau contrôle
          </button>
        )}
      </div>

      {/* ─── Stats cards ───────────────────────────────────────────── */}
      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Total</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.aPlanifier}</span>
          <span className="stat-label">À planifier</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.planifies}</span>
          <span className="stat-label">Planifiés</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.enCours}</span>
          <span className="stat-label">En cours</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.realises}</span>
          <span className="stat-label">Réalisés</span>
        </div>
        {/* Taux d'évolution global : agrégat de toutes les évaluations
            effectuées vs attendues sur la liste filtrée. Donne une vision
            macro instantanée de l'avancement collectif. */}
        <div className="stat-card resolu">
          <span
            className="stat-value"
            title={`${stats.totalRealized} évaluation(s) effectuée(s) sur ${stats.totalExpected} attendue(s)`}
          >
            {stats.tauxGlobal}%
          </span>
          <span className="stat-label">Taux d'évolution global</span>
        </div>
      </div>

      {/* ─── Barre de filtres ──────────────────────────────────────── */}
      <div className="filters-bar">
        <div className="filter-field" style={{ flex: 1, minWidth: 220 }}>
          <label>Recherche</label>
          <input
            type="text"
            placeholder="Libellé, objectif, KPI..."
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
          />
        </div>
        <div className="filter-field">
          <label>Catégorie</label>
          <select value={filters.categorie} onChange={e => updateFilter('categorie', e.target.value)}>
            <option value="">Toutes</option>
            {PLAN_CONTROLE_CATEGORIES.map(c => (
              <option key={c} value={c}>{c}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Fréquence</label>
          <select value={filters.frequence} onChange={e => updateFilter('frequence', e.target.value as FilterState['frequence'])}>
            <option value="">Toutes</option>
            {FREQUENCE_OPTIONS.map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Statut</label>
          <select value={filters.statut} onChange={e => updateFilter('statut', e.target.value as FilterState['statut'])}>
            <option value="">Tous</option>
            {STATUS_OPTIONS.map(s => (
              <option key={s} value={s}>{s}</option>
            ))}
          </select>
        </div>
        <div className="filter-field">
          <label>Responsable</label>
          <input
            type="text"
            placeholder="Nom..."
            value={filters.responsable}
            onChange={e => updateFilter('responsable', e.target.value)}
          />
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

      {/* ─── Tableau ────────────────────────────────────────────────── */}
      {loading ? (
        <p className="loading-text" style={{ marginTop: 16 }}>Chargement des contrôles...</p>
      ) : filtered.length === 0 ? (
        <p className="loading-text" style={{ marginTop: 16 }}>
          {controles.length === 0
            ? 'Aucun contrôle planifié pour le moment.'
            : 'Aucun contrôle ne correspond aux critères.'}
        </p>
      ) : (
        <div className="table-wrapper" style={{ marginTop: 16 }}>
          <table className="data-table anomalies-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Catégorie</th>
                <th>Libellé</th>
                <th>Fréquence</th>
                <th>Responsable</th>
                <th>Année</th>
                <th>Statut</th>
                <th>Taux d'évolution</th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pagedControles.map((c, i) => {
                // Nb d'évaluations effectuées pour ce contrôle (compteur préchargé).
                const evalCount = evaluationCounts.get(Number(c.id)) ?? 0
                const expected = getExpectedEvaluationsPerYear(c.frequence)
                const pct = computeEvaluationProgress(c.frequence, evalCount)
                return (
                <tr
                  key={c.id}
                  className="row-clickable"
                  onClick={() => setSelected(c)}
                  role="button"
                  tabIndex={0}
                  onKeyDown={e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setSelected(c) } }}
                >
                  <td>{pagination.start + i + 1}</td>
                  <td style={{ maxWidth: 200 }}>{c.categorie}</td>
                  <td style={{ maxWidth: 280 }}>{c.libelle}</td>
                  <td>{c.frequence}</td>
                  <td>{c.responsable || '—'}</td>
                  <td>{c.annee}</td>
                  <td>
                    <span className={`manager-pill ${getControleStatusClass(c.statut)}`}>{c.statut}</span>
                  </td>
                  <td style={{ minWidth: 130 }}>
                    {/* Taux = évaluations effectuées / attendues sur l'année (par fréquence).
                        Couleur verte forcée pour toutes les lignes — la coloration
                        automatique par seuil n'est pas pertinente ici (toute
                        progression vers l'objectif est positive). */}
                    <ProgressBar
                      value={pct}
                      color="#10b981"
                      label={`${evalCount}/${expected}`}
                      title={`${pct}% — ${evalCount} évaluation(s) sur ${expected} attendue(s) (${c.frequence.toLowerCase()})`}
                    />
                  </td>
                  {/* La cellule Actions stoppe la propagation du clic pour
                      éviter d'ouvrir le détail quand on clique sur un bouton. */}
                  <td onClick={e => e.stopPropagation()}>
                    <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
                      <button type="button" className="btn-cta btn-cta-affect" onClick={() => openEvaluation(c)}>
                        Évaluation
                      </button>
                      {/* Bouton "Affectation" : réservé aux managers (canManage).
                          Un Controleur consulte ses contrôles mais ne peut pas
                          réassigner la personne responsable. */}
                      {canManage && (
                        <button type="button" className="btn-cta btn-cta-affect" onClick={() => openAffectation(c)}>
                          Affectation
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
                )
              })}
            </tbody>
          </table>
          <Pagination
            state={pagination}
            total={filtered.length}
            itemLabel="contrôles"
          />
        </div>
      )}

      {/* ─── Modale détail ──────────────────────────────────────────── */}
      {selected && (
        <ControleDetailModal
          entry={selected}
          onClose={() => setSelected(null)}
          onEdit={canManage ? () => openEdit(selected) : undefined}
        />
      )}

      {/* ─── Modale ÉVALUATION ─────────────────────────────────────── */}
      {/* Permet de créer une évaluation pour un contrôle.
            - Le picker de période s'adapte à la fréquence du contrôle
              (jour / semaine / mois / année).
            - L'historique des évaluations existantes est listé en bas
              pour donner du contexte au manager. */}
      {evalTarget && (
        <EvaluationModal
          target={evalTarget}
          history={evalHistory}
          form={evalForm}
          setForm={setEvalForm}
          attachment={evalAttachment}
          setAttachment={setEvalAttachment}
          saving={evalSaving}
          error={evalError}
          onClose={closeEvaluation}
          onSubmit={submitEvaluation}
        />
      )}

      {/* ─── Modale AFFECTATION ────────────────────────────────────── */}
      {/* Workflow rapide pour changer le responsable d'un contrôle sans
          rouvrir le formulaire d'édition complet. Réservé aux managers
          (le bouton qui ouvre cette modale n'est rendu que si canManage).
          Le picker est pré-rempli avec la personne déjà affectée pour
          éviter de saisir un changement par erreur. */}
      {affectTarget && (
        <div className="modal-overlay" onClick={closeAffectation}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(520px, 100%)' }}>
            <div className="modal-header">
              <h2>Affecter un responsable</h2>
              <button className="modal-close" onClick={closeAffectation} aria-label="Fermer">&times;</button>
            </div>
            <div className="modal-body">
              {/* Rappel du contrôle ciblé (lecture seule) */}
              <dl className="detail-grid" style={{ marginBottom: 12 }}>
                <dt>Contrôle</dt><dd><strong>{affectTarget.libelle}</strong></dd>
                <dt>Catégorie</dt><dd>{affectTarget.categorie || '—'}</dd>
                <dt>Responsable actuel</dt>
                <dd>
                  {affectTarget.responsable
                    ? <>{affectTarget.responsable}<span style={{ color: '#888', fontSize: 12 }}> ({affectTarget.responsableEmail})</span></>
                    : '—'}
                </dd>
              </dl>

              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="affect-user">Nouveau responsable *</label>
                <UserPicker
                  id="affect-user"
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

      {/* ─── Modale création ───────────────────────────────────────── */}
      {showForm && (
        <div className="modal-overlay" onClick={closeForm}>
          <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(720px, 100%)' }}>
            <div className="modal-header">
              <h2>{editingId ? 'Modifier le contrôle' : 'Nouveau contrôle'}</h2>
              <button className="modal-close" onClick={closeForm}>&times;</button>
            </div>
            <div className="modal-body">
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="ctrl-libelle">Libellé du contrôle *</label>
                <input
                  id="ctrl-libelle"
                  type="text"
                  value={form.libelle}
                  onChange={e => updateForm('libelle', e.target.value)}
                  placeholder="Ex: Evaluation du contrôle des opérations SYSTAC"
                />
              </div>
              <div className="form-field" style={{ marginBottom: 8 }}>
                <label htmlFor="ctrl-categorie">Catégorie *</label>
                <select
                  id="ctrl-categorie"
                  value={form.categorie}
                  onChange={e => updateForm('categorie', e.target.value)}
                >
                  {PLAN_CONTROLE_CATEGORIES.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>

              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                <div className="form-field">
                  <label htmlFor="ctrl-freq">Fréquence</label>
                  <select
                    id="ctrl-freq"
                    value={form.frequence}
                    onChange={e => updateForm('frequence', e.target.value as ControleFrequence)}
                  >
                    {FREQUENCE_OPTIONS.map(f => (
                      <option key={f} value={f}>{f}</option>
                    ))}
                  </select>
                </div>
                <div className="form-field">
                  <label htmlFor="ctrl-annee">Année</label>
                  <input
                    id="ctrl-annee"
                    type="number"
                    value={form.annee}
                    onChange={e => updateForm('annee', Number(e.target.value))}
                  />
                </div>
                <div className="form-field">
                  <label htmlFor="ctrl-resp">Responsable</label>
                  {/* Champ Personne SharePoint → sélecteur Office 365 */}
                  <UserPicker
                    id="ctrl-resp"
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
                  <label htmlFor="ctrl-statut">Statut initial</label>
                  <select
                    id="ctrl-statut"
                    value={form.statut}
                    onChange={e => updateForm('statut', e.target.value as ControleStatus)}
                  >
                    {STATUS_OPTIONS.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="ctrl-obj">Objectif</label>
                <textarea
                  id="ctrl-obj"
                  rows={2}
                  value={form.objectif}
                  onChange={e => updateForm('objectif', e.target.value)}
                />
              </div>
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="ctrl-kpi">Objectif chiffré (KPI)</label>
                <textarea
                  id="ctrl-kpi"
                  rows={2}
                  value={form.objectifChiffre}
                  onChange={e => updateForm('objectifChiffre', e.target.value)}
                  placeholder="Ex: 01 rapport de contrôle mensuel"
                />
              </div>

              {/* Pièce jointe optionnelle — uploadée après création via Power
                  Automate (workflow PLAN_CONTROLE_ATTACHMENT_API_URL). Attache
                  le fichier à l'item SharePoint nouvellement créé. */}
              <div className="form-field" style={{ marginTop: 8 }}>
                <label htmlFor="ctrl-attachment">Pièce jointe (optionnel)</label>
                <input
                  id="ctrl-attachment"
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
                    : (editingId ? 'Enregistrer' : 'Créer le contrôle')}
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
 * MODALE DÉTAIL (lecture seule + bouton Modifier + pièces jointes)
 *
 * Le bouton Modifier n'apparaît que si `onEdit` est fourni — décision prise
 * côté parent en fonction du rôle (canManage).
 * ══════════════════════════════════════════════════════════════════════════ */

interface ControleDetailModalProps {
  entry: ControleEntry
  onClose: () => void
  /** Callback pour basculer en mode édition (undefined = bouton masqué). */
  onEdit?: () => void
}

function ControleDetailModal({ entry, onClose, onEdit }: ControleDetailModalProps) {
  // Historique des évaluations du contrôle : chargé une fois à l'ouverture
  // de la modale (filtre serveur OData sur plan_controle_id). Non-bloquant :
  // pendant le chargement, on affiche un placeholder.
  const [history, setHistory] = useState<ControleEvaluation[]>([])
  const [historyLoading, setHistoryLoading] = useState(true)

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  useEffect(() => {
    // Pas de setHistoryLoading(true) ici : l'état initial est déjà `true`.
    // La modale se monte à chaque ouverture (`{selected && <Modal />}`)
    // donc un useState(true) initial suffit — pas de re-fetch sur même item.
    let cancelled = false
    listEvaluationsForControle(entry.id)
      .then(data => { if (!cancelled) setHistory(data) })
      .finally(() => { if (!cancelled) setHistoryLoading(false) })
    return () => { cancelled = true }
  }, [entry.id])

  const attachments = entry.attachments ?? []

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={e => e.stopPropagation()} style={{ width: 'min(640px, 100%)' }}>
        <div className="modal-header" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <h2 style={{ flex: 1 }}>Détail du contrôle</h2>
          {onEdit && (
            <button type="button" className="btn-cta btn-cta-detail" onClick={onEdit}>
              ✎ Modifier
            </button>
          )}
          <button className="modal-close" onClick={onClose}>&times;</button>
        </div>
        <div className="modal-body">
          <dl className="detail-grid">
            <dt>Libellé</dt><dd><strong>{entry.libelle}</strong></dd>
            <dt>Catégorie</dt><dd>{entry.categorie}</dd>
            <dt>Fréquence</dt><dd>{entry.frequence}</dd>
            <dt>Responsable</dt><dd>{entry.responsable || '—'}</dd>
            <dt>Année</dt><dd>{entry.annee}</dd>
            <dt>Statut</dt>
            <dd>
              <span className={`manager-pill ${getControleStatusClass(entry.statut)}`}>{entry.statut}</span>
            </dd>
            <dt>Objectif</dt><dd>{entry.objectif || '—'}</dd>
            <dt>Objectif chiffré</dt><dd>{entry.objectifChiffre || '—'}</dd>
          </dl>

          {/* ─── Cycle de vie ────────────────────────────────────────────
              Synthèse chronologique des événements clés du contrôle :
                1. Création (entry.createdAt)
                2. Une étape par évaluation enregistrée (history)
                3. Réalisation si statut = 'Réalisé' (entry.updatedAt)
              On ne montre l'étape de réalisation que pour ce statut final
              car les autres états (À planifier / Planifié / En cours) ne
              correspondent pas à un événement daté distinct.
              Visuel : timeline verticale ; pastille colorée par type
              d'étape (création=bleu, évaluation=violet, réalisation=vert). */}
          <div style={{ marginTop: 20 }}>
            <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>Cycle de vie</h3>
            {(() => {
              type Step = {
                type: 'creation' | 'evaluation' | 'realisation'
                label: string
                date?: string
                description?: string
              }
              const steps: Step[] = []
              if (entry.createdAt) {
                steps.push({ type: 'creation', label: 'Création du contrôle', date: entry.createdAt })
              }
              // Trie l'historique du plus ancien au plus récent pour respecter
              // l'ordre chronologique de la timeline (l'API renvoie desc).
              const evals = [...history].sort((a, b) => (a.createdAt < b.createdAt ? -1 : 1))
              evals.forEach((ev, i) => {
                steps.push({
                  type: 'evaluation',
                  label: `Évaluation ${i + 1}${ev.periode ? ` — ${formatPeriodeLabel(ev.periode, entry.frequence)}` : ''}`,
                  date: ev.createdAt,
                  description: ev.observations,
                })
              })
              if (entry.statut === 'Réalisé' && entry.updatedAt) {
                steps.push({ type: 'realisation', label: 'Contrôle réalisé', date: entry.updatedAt })
              }
              const colorOf = (t: Step['type']) =>
                t === 'creation' ? '#1d4ed8' :
                t === 'evaluation' ? '#8b5cf6' :
                '#10b981'
              return (
                <ol style={{ listStyle: 'none', margin: 0, padding: '6px 0 6px 16px', position: 'relative', borderLeft: '2px solid #e5e7eb' }}>
                  {steps.map((step, i) => (
                    <li key={i} style={{ position: 'relative', padding: '4px 0 14px 14px' }}>
                      {/* Pastille colorée alignée sur la ligne verticale */}
                      <span
                        aria-hidden="true"
                        style={{
                          position: 'absolute',
                          left: -22,
                          top: 8,
                          width: 12,
                          height: 12,
                          borderRadius: '50%',
                          background: colorOf(step.type),
                          border: '2px solid #fff',
                          boxShadow: `0 0 0 2px ${colorOf(step.type)}`,
                        }}
                      />
                      <div style={{ fontSize: 13, fontWeight: 700, color: '#1a1a2e' }}>{step.label}</div>
                      {step.date && (
                        <div style={{ fontSize: 11, color: '#888', marginTop: 1 }}>
                          {new Date(step.date).toLocaleString('fr-FR')}
                        </div>
                      )}
                      {step.description && (
                        <div style={{ fontSize: 12, color: '#555', marginTop: 2, whiteSpace: 'pre-wrap' }}>
                          {step.description}
                        </div>
                      )}
                    </li>
                  ))}
                </ol>
              )
            })()}
          </div>

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

          {/* Historique des évaluations — réutilise le format chip-card de
              la modale d'évaluation (même rendu visuel pour la cohérence). */}
          <div style={{ marginTop: 20 }}>
            <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>
              Historique des évaluations ({history.length})
            </h3>
            {historyLoading ? (
              <p className="loading-text">Chargement de l'historique...</p>
            ) : history.length === 0 ? (
              <p className="loading-text">Aucune évaluation enregistrée pour ce contrôle.</p>
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
                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, fontSize: 12 }}>
                      <strong>{formatPeriodeLabel(ev.periode, entry.frequence)}</strong>
                      {ev.createdAt && (
                        <span style={{ color: '#888' }}>
                          {new Date(ev.createdAt).toLocaleDateString('fr-FR')}
                        </span>
                      )}
                    </div>
                    {ev.observations && (
                      <div style={{ marginTop: 4, fontSize: 13, color: '#444', whiteSpace: 'pre-wrap' }}>
                        {ev.observations}
                      </div>
                    )}
                    {/* Pièces jointes éventuelles de l'évaluation */}
                    {ev.attachments && ev.attachments.length > 0 && (
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


/* ══════════════════════════════════════════════════════════════════════════
 * MODALE ÉVALUATION
 *
 * Affiche en haut le formulaire de création d'une nouvelle évaluation
 * (avec un picker de période adapté à la fréquence du contrôle parent),
 * et en bas la liste des évaluations existantes pour ce contrôle.
 *
 * Pattern de période :
 *   - Quotidienne  → <input type="date">  → 'YYYY-MM-DD'
 *   - Hebdomadaire → <input type="week">  → 'YYYY-Www' (ISO week)
 *   - Mensuelle    → <input type="month"> → 'YYYY-MM'
 *   - Annuelle     → <input type="number"> → 'YYYY'
 *
 * On stocke la valeur brute du picker pour un round-trip sans ambiguïté.
 * ══════════════════════════════════════════════════════════════════════════ */

interface EvaluationModalProps {
  target: ControleEntry
  history: ControleEvaluation[]
  form: { periode: string; observations: string }
  setForm: (next: { periode: string; observations: string }) => void
  /** Pièce jointe optionnelle (uploadée après création de l'évaluation). */
  attachment: File | null
  setAttachment: (file: File | null) => void
  saving: boolean
  error: string | null
  onClose: () => void
  onSubmit: () => void
}

function EvaluationModal({
  target,
  history,
  form,
  setForm,
  attachment,
  setAttachment,
  saving,
  error,
  onClose,
  onSubmit,
}: EvaluationModalProps) {
  // Type d'input HTML pour le picker de période, déterminé une fois par
  // la fréquence du contrôle parent (immutable pendant la session modale).
  const periodeInputType = getPeriodeInputType(target.frequence)

  // Escape ferme la modale (cohérence avec les autres modales du projet)
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  // Bornes raisonnables pour le picker de période :
  //   - Pour 'number' (Annuelle), on cadre entre 2020 et 2099
  //   - Pour les autres, on laisse libre (les bornes natives suffisent)
  const yearMin = 2020
  const yearMax = 2099

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="modal"
        onClick={e => e.stopPropagation()}
        style={{ width: 'min(640px, 100%)' }}
        role="dialog"
        aria-modal="true"
      >
        <div className="modal-header">
          <h2>Évaluation du contrôle</h2>
          <button className="modal-close" onClick={onClose} aria-label="Fermer">&times;</button>
        </div>
        <div className="modal-body">
          {/* Rappel du contrôle ciblé (lecture seule) */}
          <dl className="detail-grid" style={{ marginBottom: 12 }}>
            <dt>Libellé</dt><dd><strong>{target.libelle}</strong></dd>
            <dt>Catégorie</dt><dd>{target.categorie || '—'}</dd>
            <dt>Fréquence</dt><dd>{target.frequence}</dd>
            <dt>Responsable</dt><dd>{target.responsable || '—'}</dd>
          </dl>

          {/* Formulaire de nouvelle évaluation */}
          <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>Nouvelle évaluation</h3>

          {/* Picker de période — type d'input adapté à la fréquence */}
          <div className="form-field" style={{ marginBottom: 8 }}>
            <label htmlFor="eval-periode">
              Période *
              <span style={{ fontSize: 11, color: '#888', marginLeft: 6 }}>
                ({target.frequence})
              </span>
            </label>
            {periodeInputType === 'number' ? (
              <input
                id="eval-periode"
                type="number"
                min={yearMin}
                max={yearMax}
                step={1}
                value={form.periode}
                placeholder="Ex: 2026"
                onChange={e => setForm({ ...form, periode: e.target.value })}
                disabled={saving}
              />
            ) : (
              <input
                id="eval-periode"
                type={periodeInputType}
                value={form.periode}
                onChange={e => setForm({ ...form, periode: e.target.value })}
                disabled={saving}
              />
            )}
          </div>

          {/* Observations */}
          <div className="form-field" style={{ marginBottom: 8 }}>
            <label htmlFor="eval-obs">Observations *</label>
            <textarea
              id="eval-obs"
              rows={4}
              value={form.observations}
              onChange={e => setForm({ ...form, observations: e.target.value })}
              placeholder="Constat, anomalies relevées, points d'attention..."
              disabled={saving}
            />
          </div>

          {/* Pièce jointe optionnelle — uploadée après création via Power
              Automate (workflow EVALUATION_ATTACHMENT_API_URL). L'URL renvoyée
              est ensuite ajoutée au champ urlPieceJointe de l'évaluation. */}
          <div className="form-field" style={{ marginBottom: 8 }}>
            <label htmlFor="eval-attachment">Pièce jointe (optionnel)</label>
            <input
              id="eval-attachment"
              type="file"
              onChange={e => setAttachment(e.target.files?.[0] ?? null)}
              disabled={saving}
            />
            {attachment && (
              <span className="selected-email">📎 {attachment.name}</span>
            )}
          </div>

          {error && (
            <p style={{ color: '#c0392b', fontSize: 13, margin: '8px 0' }} role="alert">
              {error}
            </p>
          )}

          <div
            className="modal-actions"
            style={{ marginTop: 8, display: 'flex', gap: 8, justifyContent: 'flex-end' }}
          >
            <button
              type="button"
              className="btn-cta btn-cta-detail"
              onClick={onClose}
              disabled={saving}
            >
              Fermer
            </button>
            <button
              type="button"
              className="btn-cta btn-cta-affect"
              onClick={onSubmit}
              disabled={saving}
            >
              {saving ? 'Enregistrement...' : "Enregistrer l'évaluation"}
            </button>
          </div>

          {/* Historique des évaluations existantes du contrôle */}
          <div style={{ marginTop: 20 }}>
            <h3 style={{ margin: '0 0 8px', fontSize: 14 }}>
              Historique des évaluations ({history.length})
            </h3>
            {history.length === 0 ? (
              <p className="loading-text">Aucune évaluation enregistrée pour le moment.</p>
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
                    <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8, fontSize: 12 }}>
                      <strong>{formatPeriodeLabel(ev.periode, target.frequence)}</strong>
                      {ev.createdAt && (
                        <span style={{ color: '#888' }}>
                          {new Date(ev.createdAt).toLocaleDateString('fr-FR')}
                        </span>
                      )}
                    </div>
                    {ev.observations && (
                      <div style={{ marginTop: 4, fontSize: 13, color: '#444', whiteSpace: 'pre-wrap' }}>
                        {ev.observations}
                      </div>
                    )}
                    {/* PJ de l'évaluation : icône + lien externe par fichier */}
                    {ev.attachments && ev.attachments.length > 0 && (
                      <ul style={{ listStyle: 'none', padding: 0, margin: '6px 0 0', display: 'flex', flexWrap: 'wrap', gap: 8 }}>
                        {ev.attachments.map((att, i) => {
                          const iconType = getAttachmentIconType(att.name)
                          return (
                            <li key={i} style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                              <span aria-hidden="true">{getAttachmentIcon(iconType)}</span>
                              <a
                                href={att.url}
                                target="_blank"
                                rel="noreferrer"
                                style={{ color: '#1d4ed8', fontSize: 12, wordBreak: 'break-all' }}
                              >
                                {att.name}
                              </a>
                            </li>
                          )
                        })}
                      </ul>
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
