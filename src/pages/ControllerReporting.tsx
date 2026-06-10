/**
 * ============================================================================
 * MODULE 3 — RAPPORT QUOTIDIEN DU CONTRÔLEUR (saisie)
 * ============================================================================
 *
 * Page accessible aux contrôleurs et aux managers.
 * Permet de consigner les actions menées sur une journée :
 *   - Plusieurs lignes (domaine, objet, action, durée, observation)
 *   - Calcul automatique : total heures, % temps occupé, statut journée
 *   - Options avancées : statut journée forcé, anomalies détectées,
 *     lien rapport, observations globales, pièces jointes
 *
 * Sauvegarde brouillon automatique :
 *   - Toutes les modifications sont persistées en localStorage avec un
 *     debounce de 600 ms (évite les écritures à chaque frappe)
 *   - Au montage : si un brouillon existe pour cet email, il est restauré
 *   - À la soumission réussie : le brouillon est effacé
 *
 * Persistance définitive :
 *   - Via activityService (localStorage actuellement, SharePoint à terme)
 *   - createReport() crée le rapport en statut 'Soumis' (manager validera)
 *
 * Validation :
 *   - validateLines() vérifie domaine, objet, action (3-500 chars), durée (5-600 min)
 *   - Au moins une ligne requise
 *   - Affichage des erreurs sous chaque champ fautif (aria-invalid + style)
 * ============================================================================
 */

import { useEffect, useMemo, useState } from 'react'
import {
  ACTIVITY_DOMAINES,
  ACTIVITY_LINE_LIMITS,
  appendReportAttachmentUrls,
  clearDraft,
  computeTotals,
  createReport,
  loadDraft,
  makeEmptyLine,
  saveDraft,
  validateLines,
  type ActivityLine,
  type LineValidationError,
} from '../lib/activityService'
import { uploadActivityAttachment } from '../lib/ticketAttachments'
import './ControllerReporting.css'

/** Props passées par Dashboard — identité du contrôleur courant. */
interface ControllerReportingProps {
  userName?: string
  userEmail?: string
}

/** Helper : renvoie la date du jour au format ISO court (YYYY-MM-DD). */
function todayISO(): string {
  return new Date().toISOString().slice(0, 10)
}

export default function ControllerReporting({ userName, userEmail }: ControllerReportingProps) {
  // Identité — fallbacks défensifs pour l'affichage si props absentes
  const email = userEmail ?? ''
  const name = userName ?? '—'

  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS DU FORMULAIRE
   * ────────────────────────────────────────────────────────────────────── */

  /** Date du rapport (par défaut aujourd'hui). */
  const [date, setDate] = useState(todayISO())
  /** Lignes d'activités. Au moins une ligne vide à l'initialisation. */
  const [lines, setLines] = useState<ActivityLine[]>([makeEmptyLine()])
  /** Nombre d'anomalies détectées par le contrôleur ce jour. */
  const [anomaliesDetectees, setAnomaliesDetectees] = useState<number>(0)
  /** Commentaires globaux libres du contrôleur. */
  const [observationsGlobales, setObservationsGlobales] = useState('')
  /** Pièces jointes locales (uploadées au moment de la soumission). */
  const [attachments, setAttachments] = useState<File[]>([])
  /** Erreurs de validation par ligne (cf. validateLines). */
  const [errors, setErrors] = useState<LineValidationError[]>([])
  /** True pendant le POST (désactive le bouton). */
  const [submitting, setSubmitting] = useState(false)
  /** Message ok/erreur après soumission. */
  const [submitMessage, setSubmitMessage] = useState<{ type: 'ok' | 'err'; text: string } | null>(null)
  /** Indicateur visuel "Brouillon restauré" pour rassurer l'utilisateur. */
  const [draftRestored, setDraftRestored] = useState(false)


  /* ──────────────────────────────────────────────────────────────────────
   * EFFETS — RESTAURATION & SAUVEGARDE BROUILLON
   * ────────────────────────────────────────────────────────────────────── */

  /**
   * Au montage : restaure le brouillon localStorage pour ce contrôleur s'il existe.
   * Déclenche le badge "Brouillon restauré" pour transparence UX.
   */
  useEffect(() => {
    if (!email) return
    const draft = loadDraft(email)
    if (draft && draft.lines && draft.lines.length > 0) {
      setDate(draft.date ?? todayISO())
      setLines(draft.lines)
      setAnomaliesDetectees(draft.anomaliesDetectees ?? 0)
      setObservationsGlobales(draft.observationsGlobales ?? '')
      setDraftRestored(true)
    }
  }, [email])

  /**
   * Auto-sauvegarde brouillon avec debounce 600 ms.
   *
   * Pourquoi debounce ? Les changements arrivent à chaque frappe utilisateur.
   * Sans debounce, on écrirait dans localStorage à chaque caractère ; avec
   * debounce, on n'écrit qu'une fois l'utilisateur a "fini" de taper.
   *
   * Le clearTimeout dans la fonction de cleanup garantit qu'un timer
   * obsolète (avant la nouvelle frappe) est annulé.
   */
  useEffect(() => {
    if (!email) return
    const handler = setTimeout(() => {
      saveDraft({
        controleurEmail: email,
        date,
        lines,
        anomaliesDetectees,
        observationsGlobales,
        updatedAt: new Date().toISOString(),
      })
    }, 600)
    return () => clearTimeout(handler)
  }, [email, date, lines, anomaliesDetectees, observationsGlobales])


  /* ──────────────────────────────────────────────────────────────────────
   * CALCULS DÉRIVÉS & HANDLERS LIGNES
   * ────────────────────────────────────────────────────────────────────── */

  /** Recalcule totaux à chaque changement des lignes (memoized). */
  const totals = useMemo(() => computeTotals(lines), [lines])

  /** Statut journée calculé automatiquement à partir des durées des lignes. */
  const effectiveStatut = totals.statutJournee

  /** Met à jour partiellement une ligne (immutable update). */
  const updateLine = (index: number, patch: Partial<ActivityLine>) => {
    setLines(prev => prev.map((l, i) => (i === index ? { ...l, ...patch } : l)))
  }

  /** Ajoute une ligne vide à la fin. */
  const addLine = () => setLines(prev => [...prev, makeEmptyLine()])

  /** Supprime la ligne à l'index donné. */
  const removeLine = (index: number) => setLines(prev => prev.filter((_, i) => i !== index))

  /** Ajoute les fichiers sélectionnés à la liste existante (multi-select). */
  const handleAttachmentChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files ?? [])
    setAttachments(prev => [...prev, ...files])
  }
  /** Retire un fichier de la liste de pièces jointes. */
  const removeAttachment = (index: number) => setAttachments(prev => prev.filter((_, i) => i !== index))

  /**
   * Réinitialise le formulaire avec confirmation.
   * Efface aussi le brouillon localStorage (impact persistant).
   */
  const handleReset = () => {
    if (!confirm('Réinitialiser le formulaire ? Le brouillon sera supprimé.')) return
    setDate(todayISO())
    setLines([makeEmptyLine()])
    setAnomaliesDetectees(0)
    setObservationsGlobales('')
    setAttachments([])
    setErrors([])
    setSubmitMessage(null)
    setDraftRestored(false)
    if (email) clearDraft(email)
  }

  /**
   * Soumission du rapport.
   *
   * Étapes :
   *   1. Vérifier email présent
   *   2. Valider toutes les lignes (validateLines) — affiche les erreurs si KO
   *   3. Préparer les métadonnées des pièces jointes
   *      ⚠ Les fichiers ne sont PAS réellement uploadés car la liste
   *      SharePoint ACTIVITE_Controleurs n'est pas encore une datasource Power Apps.
   *      On stocke seulement les NOMS pour mémoire — quand la liste sera branchée,
   *      remplacer par un upload réel + récupération des URLs.
   *   4. Appeler createReport (persistance localStorage actuellement)
   *   5. Effacer le brouillon (succès = rapport finalisé)
   *   6. Reset du formulaire
   *   7. Afficher message de succès
   */
  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setSubmitMessage(null)
    if (!email) {
      setSubmitMessage({ type: 'err', text: 'Email contrôleur introuvable.' })
      return
    }
    const validation = validateLines(lines)
    setErrors(validation)
    if (validation.length > 0) {
      setSubmitMessage({ type: 'err', text: 'Le formulaire contient des erreurs. Corrigez-les avant envoi.' })
      return
    }
    setSubmitting(true)
    try {
      // 1. Création du rapport SharePoint
      const created = await createReport({
        controleurEmail: email,
        controleurName: name,
        date,
        lines,
        anomaliesDetectees,
        observationsGlobales: observationsGlobales.trim() || undefined,
      })

      // 2. Upload des pièces jointes via le workflow Power Automate
      //    (ACTIVITY_ATTACHMENT_API_URL → liste DCPO_ACTIVICTE_CONTROLLER)
      //    Le workflow attache le fichier nativement à l'item ET retourne
      //    l'URL absolue, qu'on accumule pour la persister dans urlPieceJointes.
      let uploadFailures = 0
      const uploadedUrls: string[] = []
      if (attachments.length > 0 && created?.id) {
        for (const file of attachments) {
          try {
            const url = await uploadActivityAttachment(created.id, file)
            if (url) uploadedUrls.push(url)
          } catch (uploadErr) {
            console.error(`Échec upload pièce jointe ${file.name}`, uploadErr)
            uploadFailures++
          }
        }

        // 3. Persiste les URLs uploadées dans le champ urlPieceJointes
        //    (concaténées par " | " via appendUrl, pour cohérence avec les
        //    anomalies — permet d'afficher les liens directs côté lecture).
        if (uploadedUrls.length > 0) {
          try {
            await appendReportAttachmentUrls(created.id, uploadedUrls)
          } catch (err) {
            console.error('Échec mise à jour urlPieceJointes', err)
          }
        }
      }

      clearDraft(email)
      setSubmitMessage({
        type: uploadFailures > 0 ? 'err' : 'ok',
        text: uploadFailures > 0
          ? `Rapport soumis, mais ${uploadFailures} pièce(s) jointe(s) ont échoué à l'upload.`
          : 'Rapport soumis avec succès.',
      })
      setDate(todayISO())
      setLines([makeEmptyLine()])
      setAnomaliesDetectees(0)
      setObservationsGlobales('')
      setAttachments([])
      setErrors([])
      setDraftRestored(false)
    } catch (err) {
      console.error('Erreur soumission reporting', err)
      setSubmitMessage({ type: 'err', text: 'Erreur lors de la soumission.' })
    } finally {
      setSubmitting(false)
    }
  }

  /**
   * Recherche l'erreur de validation pour une ligne + un champ donné.
   * Utilisé dans le JSX pour afficher l'erreur SOUS l'input fautif et
   * appliquer aria-invalid="true" pour l'accessibilité.
   */
  const errorFor = (index: number, field: keyof ActivityLine): string | undefined =>
    errors.find(e => e.index === index && e.field === field)?.message

  /* ════════════════════════════════════════════════════════════════════════
   * RENDU JSX
   *
   * Layout :
   *   - Header : titre + badge "Brouillon restauré" + bouton Réinitialiser
   *   - Form :
   *     - Top row : Contrôleur (readonly) + Date du rapport
   *     - Section "Actions de la journée" : tableau éditable + bouton Ajouter
   *     - Totaux : 3 cartes (heures / temps occupé / statut journée)
   *     - Section "Options avancées" : statut forcé, anomalies, lien, observations, PJ
   *     - Message de retour (ok/err)
   *     - Bouton Soumettre
   * ════════════════════════════════════════════════════════════════════════ */
  return (
    <>
      <div className="content-header">
        <h2>Reporting contrôleur — Rapport quotidien</h2>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          {draftRestored && <span className="reporting-draft-tag">Brouillon restauré</span>}
          <button type="button" className="btn-reset-filters" onClick={handleReset}>
            Réinitialiser
          </button>
        </div>
      </div>

      <form className="reporting-form" onSubmit={handleSubmit}>
        <div className="reporting-toprow">
          <div className="form-field">
            <label htmlFor="reporting-controleur">Contrôleur</label>
            <input id="reporting-controleur" type="text" value={name} readOnly />
            {email && <span className="selected-email">{email}</span>}
          </div>
          <div className="form-field">
            <label htmlFor="reporting-date">Date du rapport</label>
            <input
              id="reporting-date"
              type="date"
              value={date}
              max={todayISO()}
              onChange={e => setDate(e.target.value)}
              required
            />
          </div>
        </div>

        <div className="reporting-section">
          <div className="reporting-section-header">
            <h3>Actions de la journée</h3>
            <button type="button" className="btn-add" onClick={addLine}>+ Ajouter une ligne</button>
          </div>

          <div className="reporting-table-wrapper">
            <table className="reporting-table">
              <thead>
                <tr>
                  <th style={{ width: 36 }}>#</th>
                  <th style={{ width: 140 }}>Domaine</th>
                  <th style={{ width: 200 }}>Objet</th>
                  <th>Action réalisée</th>
                  <th style={{ width: 110 }}>Durée (min)</th>
                  <th>Observation</th>
                  <th style={{ width: 60 }} aria-label="Suppression"></th>
                </tr>
              </thead>
              <tbody>
                {lines.map((line, index) => (
                  <tr key={line.id}>
                    <td className="row-num">{index + 1}</td>
                    <td>
                      <select
                        value={line.domaine}
                        aria-label={`Domaine ligne ${index + 1}`}
                        aria-invalid={!!errorFor(index, 'domaine')}
                        onChange={e => updateLine(index, { domaine: e.target.value })}
                      >
                        <option value="">— Choisir —</option>
                        {ACTIVITY_DOMAINES.map(d => <option key={d} value={d}>{d}</option>)}
                      </select>
                      {errorFor(index, 'domaine') && <small className="field-error">{errorFor(index, 'domaine')}</small>}
                    </td>
                    <td>
                      <input
                        type="text"
                        value={line.objet}
                        aria-label={`Objet ligne ${index + 1}`}
                        aria-invalid={!!errorFor(index, 'objet')}
                        onChange={e => updateLine(index, { objet: e.target.value })}
                        placeholder="Ex : Anomalie #T-12, Visite agence..."
                      />
                      {errorFor(index, 'objet') && <small className="field-error">{errorFor(index, 'objet')}</small>}
                    </td>
                    <td>
                      <textarea
                        rows={2}
                        value={line.action}
                        aria-label={`Action ligne ${index + 1}`}
                        aria-invalid={!!errorFor(index, 'action')}
                        maxLength={ACTIVITY_LINE_LIMITS.actionMaxLength}
                        onChange={e => updateLine(index, { action: e.target.value })}
                        placeholder="Description de l'action..."
                      />
                      <small className="field-hint">{(line.action ?? '').length}/{ACTIVITY_LINE_LIMITS.actionMaxLength}</small>
                      {errorFor(index, 'action') && <small className="field-error">{errorFor(index, 'action')}</small>}
                    </td>
                    <td>
                      <input
                        type="number"
                        min={ACTIVITY_LINE_LIMITS.dureeMin}
                        max={ACTIVITY_LINE_LIMITS.dureeMax}
                        step={5}
                        value={line.duree}
                        aria-label={`Durée ligne ${index + 1} en minutes`}
                        aria-invalid={!!errorFor(index, 'duree')}
                        onChange={e => updateLine(index, { duree: Number(e.target.value) })}
                      />
                      {errorFor(index, 'duree') && <small className="field-error">{errorFor(index, 'duree')}</small>}
                    </td>
                    <td>
                      <input
                        type="text"
                        value={line.observation ?? ''}
                        aria-label={`Observation ligne ${index + 1}`}
                        onChange={e => updateLine(index, { observation: e.target.value })}
                      />
                    </td>
                    <td>
                      <button
                        type="button"
                        className="btn-remove-row"
                        onClick={() => removeLine(index)}
                        disabled={lines.length === 1}
                        aria-label={`Supprimer ligne ${index + 1}`}
                      >
                        ✕
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        <div className="reporting-totals">
          <div className="totals-card">
            <span className="totals-label">Total heures</span>
            <span className="totals-value">{totals.totalHeures} h</span>
          </div>
          <div className="totals-card">
            <span className="totals-label">Temps occupé</span>
            <span className="totals-value">{totals.tempsOccupe} %</span>
            <progress max={100} value={totals.tempsOccupe} />
          </div>
          <div className="totals-card">
            <span className="totals-label">Statut journée</span>
            <span className={`totals-status totals-status-${effectiveStatut.toLowerCase().replace(/\s+/g, '-').replace('è', 'e')}`}>
              {effectiveStatut}
            </span>
          </div>
        </div>

        <div className="reporting-section">
          <h3>Options avancées</h3>
          <div className="form-grid">
            <div className="form-field">
              <label htmlFor="reporting-anomalies">nombres d'anomalies détectées</label>
              <input
                id="reporting-anomalies"
                type="number"
                min={0}
                value={anomaliesDetectees}
                onChange={e => setAnomaliesDetectees(Number(e.target.value))}
              />
            </div>
            <div className="form-field" style={{ gridColumn: '1 / -1' }}>
              <label htmlFor="reporting-observations">Observations globales</label>
              <textarea
                id="reporting-observations"
                rows={3}
                value={observationsGlobales}
                onChange={e => setObservationsGlobales(e.target.value)}
              />
            </div>
            <div className="form-field" style={{ gridColumn: '1 / -1' }}>
              <label htmlFor="reporting-attachments">Pièces jointes</label>
              <input id="reporting-attachments" type="file" multiple onChange={handleAttachmentChange} />
              {attachments.length > 0 && (
                <ul className="attachment-preview-list">
                  {attachments.map((f, i) => (
                    <li key={i}>
                      <span>📎 {f.name}</span>
                      <button type="button" className="btn-remove-row" onClick={() => removeAttachment(i)} aria-label={`Retirer ${f.name}`}>✕</button>
                    </li>
                  ))}
                </ul>
              )}
              <small className="field-hint">L'upload SharePoint sera réalisé après création (ACTIVITE_Controleurs).</small>
            </div>
          </div>
        </div>

        {submitMessage && (
          <div className={`reporting-message reporting-message-${submitMessage.type}`} role={submitMessage.type === 'err' ? 'alert' : 'status'}>
            {submitMessage.text}
          </div>
        )}

        <div className="reporting-submit-row">
          <button className="btn-submit" type="submit" disabled={submitting}>
            {submitting ? 'Envoi en cours...' : 'Soumettre le rapport'}
          </button>
        </div>
      </form>
    </>
  )
}
