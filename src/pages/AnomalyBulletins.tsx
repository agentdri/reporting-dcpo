/**
 * ============================================================================
 * MODULE 2 — BULLETINS D'ANOMALIES
 * ============================================================================
 *
 * Vue dédiée aux anomalies CLÔTURÉES (statut Resolu ou Clos).
 * Chaque anomalie close devient automatiquement un "bulletin consolidé"
 * affiché en card dans une grille, avec une fiche détaillée accessible
 * en modale.
 *
 * Fonctionnalités :
 *   - Récupération filtrée serveur : `field_10 eq 'Resolu' or field_10 eq 'Clos'`
 *   - Filtres client (avec bouton Rechercher pour application différée) :
 *     statut, classification, criticité, agence, agent (auteur),
 *     personne affectée, période de clôture, recherche libre
 *   - Stats : compteurs résolus/clos, délai moyen, montant cumulé
 *   - Card cliquable → modale fiche bulletin avec :
 *     identification, dates clés, caractérisation, description/causes,
 *     timeline du cycle de vie, actions à mener, pièces jointes
 *   - Bouton imprimer (CSS @media print masque sidebar/header)
 *   - Bouton SharePoint → ouvrir l'item dans l'UI SharePoint native
 *
 * Toute la logique métier (mapping, filtrage, formatage) est dans
 * src/lib/anomalyBulletin.ts.
 * ============================================================================
 */

import React, { useCallback, useEffect, useMemo, useState } from 'react'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import type { DCPO_LISTE_AGENCESRead } from '../generated/models/DCPO_LISTE_AGENCESModel'
import type { DCPO_LISTE_RESEAUXRead } from '../generated/models/DCPO_LISTE_RESEAUXModel'
import { loadAgences, loadReseaux } from '../lib/spReferenceRows'
import {
  applyBulletinFilters,
  buildConsolidatedBulletin,
  EMPTY_BULLETIN_FILTERS,
  formatAmount,
  formatDate,
  formatDateTime,
  getCriticiteClass,
  getStatusClass,
  isBulletinTicket,
  type BulletinFilters,
  type BulletinStatus,
  type ConsolidatedBulletin,
} from '../lib/anomalyBulletin'
import { getAttachmentIcon, getTicketAttachments } from '../lib/ticketAttachments'
import { formatMontantCompact } from '../lib/formatters'
import { DOMAINE_ACTIVITE_OPTIONS, TYPE_SANCTION_OPTIONS } from '../lib/referentiels'
import { Pagination } from '../components/Pagination'
import { usePagination } from '../components/usePagination'
import { UserPicker } from '../components/UserPicker'
import { ExportButtons } from '../components/ExportButtons'
import { formatDateForExport } from '../lib/exporters'
import { getAllChunkedByMonth } from '../lib/sharePointPaging'
import './AnomalyBulletins.css'
import { ModalOverlay } from '../components/ModalOverlay'

// Listes fermées pour les selects de filtre
const STATUS_OPTIONS: BulletinStatus[] = ['Tous', 'Resolu', 'Clos']
const CLASSIFICATION_OPTIONS = ['Operationnel', 'Fraude', 'Commercial']
const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']

interface AnomalyBulletinsProps {
  /** Rôle effectif de l'utilisateur connecté (utilisé pour les permissions d'édition). */
  userRole?: string
  /** Email Office 365 (réservé à des extensions futures — non utilisé directement ici). */
  userEmail?: string
}

export default function AnomalyBulletins({ userRole }: AnomalyBulletinsProps = {}) {
  /* ──────────────────────────────────────────────────────────────────────
   * ÉTATS
   * ────────────────────────────────────────────────────────────────────── */

  /** Tickets bruts (anomalies Resolu/Clos) — alimenté par fetchData. */
  const [tickets, setTickets] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  /** Référentiels pour résoudre les libellés agence/réseau dans les cards. */
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  /** Filtres en cours de saisie (binding inputs). */
  const [filters, setFilters] = useState<BulletinFilters>(EMPTY_BULLETIN_FILTERS)
  /** Filtres effectivement appliqués (snapshot au clic Rechercher). */
  const [appliedFilters, setAppliedFilters] = useState<BulletinFilters>(EMPTY_BULLETIN_FILTERS)
  /** Bulletin sélectionné pour la modale fiche détaillée. */
  const [selected, setSelected] = useState<ConsolidatedBulletin | null>(null)
  /**
   * Bulletin actuellement en cours d'édition (réservé au rôle Directeur).
   * Décorrélé de `selected` pour permettre une UX claire : on passe d'une
   * vue lecture seule à une vue formulaire dédiée, sans superposer les
   * deux modales.
   */
  const [editing, setEditing] = useState<ConsolidatedBulletin | null>(null)
  /**
   * Permission d'édition rétroactive (Resolu / Clos) — ouverte aux rôles
   * MANAGERS : Directeur ET Chef_Departement. Un contrôleur n'a jamais
   * accès à cette édition (l'anomalie est déjà clôturée, il ne peut plus
   * la modifier).
   *
   * Comparaison case-insensitive et accent-insensitive pour absorber les
   * variantes de saisie SP (`Directeur`, `directeur`, `Chef_Departement`,
   * `chef_departement`, etc.).
   */
  const canManagerEdit = (() => {
    const normalized = (userRole ?? '')
      .normalize('NFD').replace(/\p{Diacritic}/gu, '')
      .trim().toLowerCase()
    return normalized === 'directeur' || normalized === 'chef_departement'
  })()

  /**
   * Récupère depuis SharePoint UNIQUEMENT les anomalies Resolu/Clos.
   *
   * ⚠ VOLUMÉTRIE : le connecteur SP plafonne pratiquement à ~500 items
   * par appel `getAll` (bug historique = seuls les 100/500 items les plus
   * récents remontaient dès que la liste dépassait ce seuil). On chunke
   * donc la plage par MOIS calendaire via `getAllChunkedByMonth` — chaque
   * chunk restant très en-dessous de 500 items dans la pratique métier.
   *
   * Plage appliquée :
   *   - Si l'utilisateur a saisi Date du / Date au (appliqués via
   *     "Rechercher"), on utilise ces bornes.
   *   - Sinon défaut = année en cours (1er janvier → 31 décembre).
   *   La plage porte sur la colonne `Created` (indexée par défaut SP).
   *
   * Le filtre `field_10 eq 'Resolu' or field_10 eq 'Clos'` est combiné
   * avec le filtre de plage — on ne récupère JAMAIS les Ouvert/En cours
   * sur cette page.
   *
   * En parallèle : agences + réseaux pour résoudre les libellés des cards.
   */
  const fetchData = useCallback(async () => {
    setLoading(true)
    setError(null)
    try {
      const currentYear = new Date().getFullYear()
      const defaultFrom = `${currentYear}-01-01`
      const defaultTo = `${currentYear}-12-31`
      const from = appliedFilters.dateFrom || defaultFrom
      const to = appliedFilters.dateTo || defaultTo

      const [ticketsData, agencesRows, reseauxRows] = await Promise.all([
        getAllChunkedByMonth<DCPO_LISTE_ANORMALIERead>(DCPO_LISTE_ANORMALIEService, {
          from,
          to,
          dateColumn: 'Created',
          extraFilter: "field_10 eq 'Resolu' or field_10 eq 'Clos'",
          orderBy: ['Created desc'],
        }),
        loadAgences(),
        loadReseaux(),
      ])
      setAgences(agencesRows)
      setReseaux(reseauxRows)
      setTickets(ticketsData.filter(isBulletinTicket))
    } catch (err) {
      console.error('Erreur chargement bulletins', err)
      setError('Impossible de charger les bulletins.')
    } finally {
      setLoading(false)
    }
  }, [appliedFilters.dateFrom, appliedFilters.dateTo])

  // Chargement initial au montage + refetch dès que la plage appliquée change.
  useEffect(() => {
    fetchData()
  }, [fetchData])

  /**
   * Mappe chaque ticket brut en bulletin consolidé (cf. anomalyBulletin.ts).
   * useMemo : recalcule UNIQUEMENT si tickets/agences/reseaux changent.
   */
  const bulletins = useMemo(
    () => tickets.map(t => buildConsolidatedBulletin(t, agences, reseaux)),
    [tickets, agences, reseaux],
  )

  /** Bulletins après application des filtres utilisateur (côté client). */
  const filtered = useMemo(() => applyBulletinFilters(bulletins, appliedFilters), [bulletins, appliedFilters])

  /**
   * Pagination — porte sur la liste FILTRÉE (les filtres réduisent le total
   * paginé). resetKey = signature JSON des filtres : un changement applique
   * automatiquement le retour en page 1 (sinon UX cassée : on serait page 5
   * d'une liste qui n'a plus que 3 pages).
   *
   * Taille de page configurable globalement via DEFAULT_PAGE_SIZE dans
   * src/components/usePagination.ts (cf. doc de paramétrage).
   */
  const pagination = usePagination({
    total: filtered.length,
    resetKey: JSON.stringify(appliedFilters),
  })
  /** Sous-ensemble visible sur la page courante (tranche start..end). */
  const pagedBulletins = useMemo(
    () => filtered.slice(pagination.start, pagination.end),
    [filtered, pagination.start, pagination.end],
  )

  /**
   * Statistiques agrégées pour les cards en haut de page :
   *   - resolus / clos : compteurs par statut
   *   - totalMontant   : somme des montants (champ field_8)
   *   - avgDelay       : délai moyen en jours, calculé uniquement sur les
   *                      bulletins ayant un délai > 0 (évite la pollution
   *                      par les valeurs absentes)
   */
  const stats = useMemo(() => {
    const resolus = filtered.filter(b => b.statut === 'Resolu').length
    const clos = filtered.filter(b => b.statut === 'Clos').length
    const totalMontant = filtered.reduce((s, b) => s + (b.montant ?? 0), 0)
    const delais = filtered.map(b => b.delayDays ?? 0).filter(n => n > 0)
    const avgDelay = delais.length ? Math.round(delais.reduce((s, n) => s + n, 0) / delais.length) : 0
    return { total: filtered.length, resolus, clos, totalMontant, avgDelay }
  }, [filtered])

  /** Helper générique pour patcher un champ de filtre. */
  const updateFilter = <K extends keyof BulletinFilters>(key: K, value: BulletinFilters[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }

  /** Au clic Rechercher : snapshot draft → applied. */
  const applyFilters = () => setAppliedFilters(filters)

  /** Réinitialise tout (saisie + applied). */
  const resetFilters = () => {
    setFilters(EMPTY_BULLETIN_FILTERS)
    setAppliedFilters(EMPTY_BULLETIN_FILTERS)
  }

  /**
   * Lance l'impression du bulletin (CSS @media print masque sidebar/header).
   * L'utilisateur peut ensuite "Enregistrer comme PDF" depuis le dialogue
   * d'impression du navigateur.
   */
  const printBulletin = () => {
    window.print()
  }

  return (
    <>
      <div className="content-header">
        <h2>Fiche récapitulatif de l'anomalie</h2>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <ExportButtons
            filename="bulletins_anomalies"
            pdfTitle="Bulletins d'anomalies"
            getHeaders={() => [
              'Numéro', 'Titre', 'Statut', 'Criticité', 'Classification',
              'Agence', 'Réseau', 'Date déclaration', 'Date clôture',
              'Délai (jours)', 'Déclarant', 'Auteur', 'Personne affectée',
              'Montant', 'Description', 'Cause immédiate', 'Cause racine',
            ]}
            getRows={() => filtered.map(b => [
              b.numero,
              b.titre,
              b.statut,
              b.criticite,
              b.classification,
              b.agenceLabel,
              b.reseauLabel,
              formatDateForExport(b.declarationDate?.toISOString()),
              formatDateForExport(b.closureDate?.toISOString()),
              b.delayDays ?? '',
              b.declarantName,
              b.auteurName,
              b.affecteName,
              b.montant ?? '',
              b.description,
              b.causeImmediate,
              b.causeRacine,
            ])}
            disabled={loading}
          />
          <button className="btn-add" onClick={fetchData} disabled={loading} type="button">
            {loading ? 'Actualisation...' : 'Actualiser'}
          </button>
        </div>
      </div>

      <div className="stats-cards">
        <div className="stat-card total">
          <span className="stat-value">{stats.total}</span>
          <span className="stat-label">Bulletins</span>
        </div>
        <div className="stat-card resolu">
          <span className="stat-value">{stats.resolus}</span>
          <span className="stat-label">Résolus</span>
        </div>
        <div className="stat-card clos">
          <span className="stat-value">{stats.clos}</span>
          <span className="stat-label">Clos</span>
        </div>
        <div className="stat-card en-cours">
          <span className="stat-value">{stats.avgDelay}<small> j</small></span>
          <span className="stat-label">Délai moyen</span>
        </div>
        <div className="stat-card montant">
          {/* Format compact (millions au-delà d'1 M) pour éviter le débordement. */}
          <span className="stat-value">{formatMontantCompact(stats.totalMontant)}</span>
          <span className="stat-label">Montant cumulé</span>
        </div>
      </div>

      <div className="bulletins-toolbar">
        <div className="bulletins-search">
          <span className="search-icon" aria-hidden="true">🔍</span>
          <input
            type="text"
            placeholder="Rechercher un bulletin (n°, titre, agence, cause, ...)"
            value={filters.search}
            onChange={e => updateFilter('search', e.target.value)}
            onKeyDown={e => { if (e.key === 'Enter') applyFilters() }}
            aria-label="Recherche"
          />
        </div>
        <div className="bulletins-filters">
          <div className="filter-field">
            <label>Statut</label>
            <select value={filters.status} onChange={e => updateFilter('status', e.target.value as BulletinStatus)}>
              {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Classification</label>
            <select value={filters.classification} onChange={e => updateFilter('classification', e.target.value)}>
              <option value="">Toutes</option>
              {CLASSIFICATION_OPTIONS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Criticité</label>
            <select value={filters.criticite} onChange={e => updateFilter('criticite', e.target.value)}>
              <option value="">Toutes</option>
              {CRITICITE_OPTIONS.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Agence</label>
            <select value={filters.agence} onChange={e => updateFilter('agence', e.target.value)}>
              <option value="">Toutes</option>
              {agences.map(a => (
                <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
              ))}
            </select>
          </div>
          <div className="filter-field">
            <label>Réseau</label>
            <select value={filters.reseau} onChange={e => updateFilter('reseau', e.target.value)}>
              <option value="">Tous</option>
              {reseaux.map(r => (
                <option key={r.ID} value={String(r.ID)}>{r.field_1 ?? r.Title}</option>
              ))}
            </select>
          </div>
          <div className="filter-field">
            <label>Domaine d'activité</label>
            <select
              value={filters.domaineActivite}
              onChange={e => updateFilter('domaineActivite', e.target.value)}
            >
              <option value="">Tous</option>
              {DOMAINE_ACTIVITE_OPTIONS.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
          <div className="filter-field">
            <label>Agent (auteur)</label>
            <input
              type="text"
              placeholder="Nom ou email..."
              value={filters.agent}
              onChange={e => updateFilter('agent', e.target.value)}
            />
          </div>
          <div className="filter-field">
            <label>Personne affectée</label>
            <input
              type="text"
              placeholder="Nom ou email..."
              value={filters.affecte}
              onChange={e => updateFilter('affecte', e.target.value)}
            />
          </div>
          <div className="filter-field">
            <label>Déclarant</label>
            <input
              type="text"
              placeholder="Nom ou email..."
              value={filters.declarant}
              onChange={e => updateFilter('declarant', e.target.value)}
            />
          </div>
          <div className="filter-field">
            <label>Date du</label>
            <input type="date" value={filters.dateFrom} onChange={e => updateFilter('dateFrom', e.target.value)} />
          </div>
          <div className="filter-field">
            <label>Date au</label>
            <input type="date" value={filters.dateTo} onChange={e => updateFilter('dateTo', e.target.value)} />
          </div>
          <button type="button" className="btn-search-filters" onClick={applyFilters} disabled={loading}>
            Rechercher
          </button>
          <button type="button" className="btn-reset-filters" onClick={resetFilters}>Réinitialiser</button>
        </div>
      </div>

      {error && <div className="bulletin-error" role="alert">{error}</div>}

      {loading ? (
        <p className="loading-text">Chargement des bulletins...</p>
      ) : filtered.length === 0 ? (
        <div className="bulletins-empty">
          <p>Aucun bulletin ne correspond aux critères. {tickets.length === 0 && 'Aucune anomalie résolue ou close trouvée.'}</p>
        </div>
      ) : (
        <>
          <div className="bulletins-grid">
            {pagedBulletins.map(b => (
              <BulletinCard key={b.ticket.ID ?? b.numero} bulletin={b} onOpen={() => setSelected(b)} />
            ))}
          </div>
          {/* Barre de pagination — taille de page configurable globalement
              via DEFAULT_PAGE_SIZE / PAGE_SIZE_OPTIONS dans usePagination.ts */}
          <Pagination
            state={pagination}
            total={filtered.length}
            itemLabel="bulletins"
          />
        </>
      )}

      {selected && !editing && (
        <BulletinModal
          bulletin={selected}
          onClose={() => setSelected(null)}
          onPrint={printBulletin}
          canEdit={canManagerEdit}
          onEdit={() => setEditing(selected)}
        />
      )}

      {editing && (
        <BulletinEditModal
          bulletin={editing}
          agences={agences}
          reseaux={reseaux}
          onClose={() => setEditing(null)}
          onSaved={async () => {
            setEditing(null)
            await fetchData()
            setSelected(null) // ferme aussi la modale lecture (les données vont être remplacées)
          }}
        />
      )}
    </>
  )
}

/**
 * Card individuelle d'un bulletin dans la grille.
 *
 * Comportements UX :
 *   - Cliquable (onClick) — ouvre la modale fiche
 *   - Accessible clavier : tabIndex={0} + onKeyDown (Enter / Space)
 *   - role="button" + aria-label pour les lecteurs d'écran
 *
 * Visuellement :
 *   - Numéro + status chip en header
 *   - Titre tronqué sur 2 lignes (CSS line-clamp)
 *   - Tags (criticité + classification)
 *   - Grille 2 colonnes avec dates clés et montant
 *   - Footer : qui est affecté + indicateur "Ouvrir →"
 */
function BulletinCard({ bulletin, onOpen }: { bulletin: ConsolidatedBulletin; onOpen: () => void }) {
  return (
    <article className="bulletin-card" tabIndex={0} onClick={onOpen} onKeyDown={e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onOpen() }
    }} role="button" aria-label={`Ouvrir bulletin ${bulletin.numero}`}>
      <header className="bulletin-card-header">
        <span className="bulletin-num">{bulletin.numero}</span>
        <span className={`bulletin-status ${getStatusClass(bulletin.statut)}`}>{bulletin.statut || '—'}</span>
      </header>
      <h3 className="bulletin-title">{bulletin.titre}</h3>
      <div className="bulletin-meta">
        <span className={`bulletin-criticite ${getCriticiteClass(bulletin.criticite)}`}>
          {bulletin.criticite || 'Criticité —'}
        </span>
        {bulletin.classification && <span className="bulletin-tag">{bulletin.classification}</span>}
      </div>
      <dl className="bulletin-card-grid">
        <div><dt>Agence</dt><dd>{bulletin.agenceLabel || '—'}</dd></div>
        <div><dt>Réseau</dt><dd>{bulletin.reseauLabel || '—'}</dd></div>
        <div><dt>Déclaré le</dt><dd>{formatDate(bulletin.declarationDate)}</dd></div>
        <div><dt>Clôturé le</dt><dd>{formatDate(bulletin.closureDate ?? bulletin.regularizationDate)}</dd></div>
        <div><dt>Délai</dt><dd>{bulletin.delayDays !== undefined ? `${bulletin.delayDays} j` : '—'}</dd></div>
        <div><dt>Montant</dt><dd>{formatAmount(bulletin.montant)}</dd></div>
      </dl>
      <footer className="bulletin-card-footer">
        <span className="bulletin-affecte">{bulletin.affecteName !== '—' ? `Affecté à ${bulletin.affecteName}` : 'Non affecté'}</span>
        <span className="bulletin-open-hint">Ouvrir →</span>
      </footer>
    </article>
  )
}

/** Props de la modale fiche bulletin. */
interface BulletinModalProps {
  bulletin: ConsolidatedBulletin
  onClose: () => void
  onPrint: () => void
  /** Si true, affiche le bouton "Modifier" (rôle Directeur). */
  canEdit?: boolean
  /** Handler du clic "Modifier" — ouvre la modale d'édition. */
  onEdit?: () => void
}

/**
 * Fiche bulletin complète en modale.
 *
 * Sections affichées :
 *   1. Header : numéro, titre, actions (Imprimer, SharePoint, Fermer)
 *   2. Status row : chips statut + criticité + classification + délai
 *   3. Identification + Dates clés (rangée 2 colonnes desktop)
 *   4. Caractérisation
 *   5. Description & causes (description + cause immédiate + cause racine
 *      + observations si présentes)
 *   6. Cycle de vie (timeline buildLifecycleSteps)
 *   7. Actions à mener (checklist disabled, cochée si statut Clos)
 *   8. Pièces jointes (avec icônes par type)
 *
 * Raccourci clavier : Escape ferme la modale (handler global window).
 */
export function BulletinModal({ bulletin, onClose, onPrint, canEdit, onEdit }: BulletinModalProps) {
  const attachments = getTicketAttachments(bulletin.ticket)

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <ModalOverlay onClose={onClose} className="bulletin-modal-overlay">
      <div
        className="modal bulletin-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="bulletin-modal-title"
      >
        <header className="bulletin-modal-header">
          <div className="bulletin-modal-title-block">
            <span className="bulletin-modal-num">{bulletin.numero}</span>
            <h2 id="bulletin-modal-title">{bulletin.titre}</h2>
          </div>
          <div className="bulletin-modal-actions no-print">
            {canEdit && onEdit && (
              <button
                type="button"
                className="btn-affect"
                onClick={onEdit}
                aria-label="Modifier l'anomalie clôturée"
              >
                ✏️ Modifier
              </button>
            )}
            <button type="button" className="btn-detail" onClick={onPrint} aria-label="Imprimer le bulletin">
              🖨️ Imprimer
            </button>
            {bulletin.sharePointUrl && (
              <a className="btn-affect btn-link" href={bulletin.sharePointUrl} target="_blank" rel="noreferrer">
                🔗 SharePoint
              </a>
            )}
            <button className="modal-close" onClick={onClose} aria-label="Fermer">&times;</button>
          </div>
        </header>

        <div className="bulletin-modal-body">
          <section className="bulletin-section bulletin-section-status">
            <span className={`bulletin-status ${getStatusClass(bulletin.statut)}`}>{bulletin.statut || '—'}</span>
            <span className={`bulletin-criticite ${getCriticiteClass(bulletin.criticite)}`}>
              {bulletin.criticite || 'Criticité —'}
            </span>
            {bulletin.classification && <span className="bulletin-tag">{bulletin.classification}</span>}
            {bulletin.delayDays !== undefined && (
              <span className="bulletin-tag bulletin-tag-delay">⏱ Délai : {bulletin.delayDays} j</span>
            )}
          </section>

          <div className="bulletin-section-row">
            <section className="bulletin-section">
              <h3>Identification</h3>
              <dl className="bulletin-detail-grid bulletin-detail-grid-2">
                <div><dt>Déclarant</dt><dd>{bulletin.declarantName}</dd></div>
                <div><dt>Auteur</dt><dd>{bulletin.auteurName}</dd></div>
                <div><dt>Personne affectée</dt><dd>{bulletin.affecteName}</dd></div>
                <div><dt>Agence</dt><dd>{bulletin.agenceLabel || '—'}</dd></div>
                <div><dt>Réseau</dt><dd>{bulletin.reseauLabel || '—'}</dd></div>
                <div><dt>Domaine</dt><dd>{bulletin.domaine || '—'}</dd></div>
              </dl>
            </section>

            <section className="bulletin-section">
              <h3>Dates clés</h3>
              <dl className="bulletin-detail-grid bulletin-detail-grid-2">
                <div><dt>Déclaration</dt><dd>{formatDate(bulletin.declarationDate)}</dd></div>
                <div><dt>Ouverture ticket</dt><dd>{formatDate(bulletin.openingDate)}</dd></div>
                <div><dt>Régularisation</dt><dd>{formatDate(bulletin.regularizationDate)}</dd></div>
                <div><dt>Clôture</dt><dd>{formatDate(bulletin.closureDate)}</dd></div>
                <div><dt>Délai</dt><dd>{bulletin.delayDays !== undefined ? `${bulletin.delayDays} j` : '—'}</dd></div>
              </dl>
            </section>
          </div>

          <section className="bulletin-section">
            <h3>Caractérisation</h3>
            <dl className="bulletin-detail-grid">
              <div><dt>Classification</dt><dd>{bulletin.classification || '—'}</dd></div>
              <div><dt>Criticité</dt><dd>{bulletin.criticite || '—'}</dd></div>
              <div><dt>Type de risque</dt><dd>{bulletin.natureRisque || '—'}</dd></div>
              <div><dt>Montant</dt><dd>{formatAmount(bulletin.montant)}</dd></div>
              <div><dt>Occurrences</dt><dd>{bulletin.occurrences ?? '—'}</dd></div>
              <div><dt>Statut</dt><dd>{bulletin.statut || '—'}</dd></div>
              {/* typeSanction : renseigné à la clôture par le contrôleur
                  via la modale de résolution. Sert au suivi disciplinaire
                  ou opérationnel des suites données à l'anomalie. */}
              <div><dt>Mode de traitement</dt><dd>{bulletin.typeSanction || '—'}</dd></div>
            </dl>
          </section>

          <section className="bulletin-section">
            <h3>Description &amp; causes</h3>
            <div className="bulletin-block">
              <h4>Description</h4>
              <p>{bulletin.description || '—'}</p>
            </div>
            <div className="bulletin-block">
              <h4>Cause immédiate</h4>
              <p>{bulletin.causeImmediate || '—'}</p>
            </div>
            <div className="bulletin-block">
              <h4>Cause racine</h4>
              <p>{bulletin.causeRacine || '—'}</p>
            </div>
            {bulletin.observations && (
              <div className="bulletin-block">
                <h4>Observations</h4>
                <p>{bulletin.observations}</p>
              </div>
            )}
          </section>

          <section className="bulletin-section">
            <h3>Cycle de vie</h3>
            {bulletin.lifecycle.length === 0 ? (
              <p className="bulletin-empty-block">Aucun événement enregistré.</p>
            ) : (
              <ol className="bulletin-timeline">
                {bulletin.lifecycle.map((step, i) => (
                  <li key={i} className={`timeline-item timeline-${step.type}`}>
                    <span className="timeline-dot" aria-hidden="true" />
                    <div className="timeline-content">
                      <div className="timeline-label">{step.label}</div>
                      {step.date && <div className="timeline-date">{formatDateTime(step.date)}</div>}
                      {step.description && <div className="timeline-desc">{step.description}</div>}
                    </div>
                  </li>
                ))}
              </ol>
            )}
          </section>

          <section className="bulletin-section">
            <h3>Actions à mener</h3>
            {bulletin.actions.length === 0 ? (
              <p className="bulletin-empty-block">Aucune action consignée.</p>
            ) : (
              <ul className="bulletin-actions-list">
                {bulletin.actions.map((action, i) => (
                  <li key={i}>
                    <input type="checkbox" id={`act-${i}`} disabled defaultChecked={bulletin.statut === 'Clos'} />
                    <label htmlFor={`act-${i}`}>{action}</label>
                  </li>
                ))}
              </ul>
            )}
          </section>

          <section className="bulletin-section">
            <h3>Pièces jointes</h3>
            {attachments.length === 0 ? (
              <p className="bulletin-empty-block">Aucune pièce jointe.</p>
            ) : (
              <ul className="bulletin-attachments">
                {attachments.map((att, i) => (
                  <li key={i}>
                    <span className="att-icon" aria-hidden="true">{getAttachmentIcon(att.iconType)}</span>
                    <a href={att.url} target="_blank" rel="noreferrer">{att.name}</a>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </div>
      </div>
    </ModalOverlay>
  )
}


/* ══════════════════════════════════════════════════════════════════════════
 * MODALE D'ÉDITION — RÉSERVÉE AU RÔLE DIRECTEUR
 *
 * Permet à un directeur de RECTIFIER une anomalie déjà clôturée (Resolu ou
 * Clos). Champs éditables limités à ceux susceptibles d'être corrigés
 * après coup (statut, dates de régularisation/clôture, description
 * consolidée, montant, mode de traitement, actions menées).
 *
 * Les champs identifiants (déclarant, agence, criticité, classification…)
 * restent NON éditables : ils sont figés métier — modifier ces champs
 * reviendrait à "réécrire l'histoire" et n'est pas attendu pour une
 * rectification post-clôture.
 *
 * Sauvegarde via DCPO_LISTE_ANORMALIEService.update + refresh complet
 * de la liste (onSaved).
 * ══════════════════════════════════════════════════════════════════════════ */

interface BulletinEditModalProps {
  bulletin: ConsolidatedBulletin
  agences: DCPO_LISTE_AGENCESRead[]
  reseaux: DCPO_LISTE_RESEAUXRead[]
  onClose: () => void
  /** Callback déclenché après une sauvegarde réussie (le parent refresh). */
  onSaved: () => void | Promise<void>
}

/** Extrait la portion `YYYY-MM-DD` d'un ISO ou `''` si vide. */
function toDateInput(value?: string | null): string {
  if (!value) return ''
  const m = value.match(/^(\d{4}-\d{2}-\d{2})/)
  return m ? m[1] : ''
}

/** Format Person SP standard (Claims). */
function toClaims(email: string): string {
  return `i:0#.f|membership|${email.toLowerCase()}`
}

/**
 * Échappe les caractères HTML spéciaux avant insertion dans le champ riche
 * field_4 (évite l'injection XSS via un texte saisi par le Directeur).
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
 * Reconstruit le champ SharePoint `field_4` (seule colonne HTML réellement
 * persistée pour la description) à partir des sections éditées.
 *
 * Pourquoi : `causeImmediate`, `causeRacine`, `observationsBulletin` et
 * `natureRisque` NE SONT PAS des colonnes réelles de la liste SharePoint
 * DCPO_LISTE_ANORMALIE (cf. .power/schemas/sharepointonline/dcpo_liste_anormalie.Schema.json
 * — absentes du schéma). Les écrire dans le payload d'update fait échouer
 * silencieusement (ou entièrement) la sauvegarde, et de toute façon la
 * lecture (buildConsolidatedBulletin) les re-dérive en priorité depuis
 * field_4 via parseDescriptionSections. On encode donc ici les sections
 * dans le MÊME format texte que celui attendu au parsing pour qu'elles
 * survivent à un aller-retour complet.
 */
function buildBulletinFieldHtml(
  description: string,
  causeImmediate: string,
  causeRacine: string,
  observations: string,
  natureRisque: string,
): string {
  const blocks: string[] = []
  if (description.trim()) {
    blocks.push(`<p>${escapeHtml(description).replace(/\n/g, '<br/>')}</p>`)
  }
  if (causeImmediate.trim()) {
    blocks.push(`<p><strong>Cause immédiate :</strong> ${escapeHtml(causeImmediate)}</p>`)
  }
  if (causeRacine.trim()) {
    blocks.push(`<p><strong>Cause racine :</strong> ${escapeHtml(causeRacine)}</p>`)
  }
  if (observations.trim()) {
    blocks.push(`<p><strong>Observations :</strong> ${escapeHtml(observations)}</p>`)
  }
  if (natureRisque.trim()) {
    blocks.push(`<p><strong>Type de risque :</strong> ${escapeHtml(natureRisque)}</p>`)
  }
  return blocks.join('')
}

const CLASSIFICATION_OPTIONS_EDIT = ['Operationnel', 'Fraude', 'Commercial']
const CRITICITE_OPTIONS_EDIT = ['Faible', 'Moyenne', 'Haute', 'Critique']

function BulletinEditModal({ bulletin, agences, reseaux, onClose, onSaved }: BulletinEditModalProps) {
  const ticket = bulletin.ticket

  // ─── Section Identification ────────────────────────────────────────────
  const [title, setTitle] = useState(ticket.Title ?? '')
  const [auteurName, setAuteurName] = useState(ticket.auteur_anormalie?.DisplayName ?? '')
  const [auteurEmail, setAuteurEmail] = useState(ticket.auteur_anormalie?.Email ?? '')
  const [affecteName, setAffecteName] = useState(ticket.personneAffecter?.DisplayName ?? '')
  const [affecteEmail, setAffecteEmail] = useState(ticket.personneAffecter?.Email ?? '')

  // ─── Section Localisation ──────────────────────────────────────────────
  const [agence, setAgence] = useState(ticket.field_6 ?? '')
  const [reseau, setReseau] = useState(ticket.field_7 ?? '')
  const [domaineActivite, setDomaineActivite] = useState(ticket.domaineActivite ?? '')

  // ─── Section Caractérisation ──────────────────────────────────────────
  const [classification, setClassification] = useState(ticket.field_5 ?? '')
  const [criticite, setCriticite] = useState(ticket.criticiteAnomalie ?? '')
  const [natureRisque, setNatureRisque] = useState(bulletin.natureRisque ?? '')
  const [montant, setMontant] = useState(
    ticket.field_8 !== undefined && ticket.field_8 !== null ? String(ticket.field_8) : '',
  )
  const [typeSanction, setTypeSanction] = useState(ticket.typeSanction ?? '')

  // ─── Section Dates & statut ───────────────────────────────────────────
  const [statut, setStatut] = useState(ticket.field_10 ?? 'Resolu')
  const [dateSurvenance, setDateSurvenance] = useState(toDateInput(ticket.field_0))
  const [dateOuverture, setDateOuverture] = useState(toDateInput(ticket.dateOuvertureTicket))
  const [dateRegul, setDateRegul] = useState(toDateInput(ticket.field_9))
  const [dateCloture, setDateCloture] = useState(toDateInput(ticket.date_cloture_ticket))

  // ─── Section Description & causes ─────────────────────────────────────
  //
  // Pré-remplissage depuis le BULLETIN CONSOLIDÉ (pas depuis extTicket) :
  // causeImmediate/causeRacine/observations/natureRisque ne sont pas des
  // colonnes SharePoint réelles (cf. buildBulletinFieldHtml ci-dessus) —
  // `bulletin.*` est la valeur correctement dérivée par parsing de field_4,
  // donc la seule source fiable pour ne pas repartir de champs vides.
  const [description, setDescription] = useState(bulletin.description ?? '')
  const [causeImmediate, setCauseImmediate] = useState(bulletin.causeImmediate ?? '')
  const [causeRacine, setCauseRacine] = useState(bulletin.causeRacine ?? '')
  const [actionsMenees, setActionsMenees] = useState(ticket.actionsMenees ?? '')
  const [observations, setObservations] = useState(bulletin.observations ?? '')

  const [saving, setSaving] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && !saving) onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose, saving])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    if (!ticket.ID) {
      setError('Identifiant de ticket introuvable.')
      return
    }
    // Règle métier : la date de clôture ne peut pas être antérieure à la
    // date d'ouverture du ticket (les deux sont éditables ici, donc on
    // compare les valeurs COURANTES du formulaire, pas celles d'origine).
    if (dateOuverture && dateCloture && dateCloture < dateOuverture) {
      setError("La date de clôture ne peut pas être antérieure à la date d'ouverture du ticket.")
      return
    }
    setSaving(true)
    setError(null)
    try {
      const montantNum = montant.trim() === '' ? undefined : Number(montant.replace(/\s/g, '').replace(',', '.'))
      if (montant.trim() !== '' && (montantNum === undefined || Number.isNaN(montantNum))) {
        throw new Error('Le montant doit être un nombre valide.')
      }

      // Payload complet : on écrit TOUS les champs (même ceux inchangés).
      // C'est cohérent avec le pattern OData PATCH — un champ absent reste
      // à sa valeur SP, un champ présent est mis à jour.
      //
      // IMPORTANT : `causeImmediate`, `causeRacine`, `observationsBulletin`
      // et `natureRisque` NE SONT PAS des colonnes de la liste SharePoint
      // (absentes de .power/schemas/.../dcpo_liste_anormalie.Schema.json).
      // Les envoyer faisait échouer la sauvegarde (silencieusement ou en
      // totalité selon le connecteur) — c'était la cause de la non-
      // persistance des modifications sur une anomalie clôturée. Ces
      // valeurs sont désormais encodées comme sections dans field_4 (seule
      // colonne HTML réellement persistée), via buildBulletinFieldHtml —
      // symétrique au parsing fait à la lecture (parseDescriptionSections).
      const payload: Record<string, unknown> = {
        Title: title,
        // Identification (Person) : encodage Claims. Un email vide efface
        // le champ Person côté SP (comportement souhaité si un manager
        // veut retirer une affectation).
        auteur_anormalie: auteurEmail
          ? {
              '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
              Claims: toClaims(auteurEmail),
            }
          : null,
        personneAffecter: affecteEmail
          ? {
              '@odata.type': '#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser',
              Claims: toClaims(affecteEmail),
            }
          : null,
        // Localisation
        field_6: agence,
        field_7: reseau,
        domaineActivite,
        // Caractérisation
        field_5: classification,
        criticiteAnomalie: criticite,
        typeSanction,
        // Statut & dates (T00:00:00Z pour cohérence UTC calendaire)
        field_10: statut,
        field_0: dateSurvenance ? `${dateSurvenance}T00:00:00Z` : '',
        dateOuvertureTicket: dateOuverture ? `${dateOuverture}T00:00:00Z` : '',
        field_9: dateRegul ? `${dateRegul}T00:00:00Z` : '',
        date_cloture_ticket: dateCloture ? `${dateCloture}T00:00:00Z` : '',
        // Description & causes — tout concentré dans field_4 (cf. note ci-dessus)
        field_4: buildBulletinFieldHtml(description, causeImmediate, causeRacine, observations, natureRisque),
        actionsMenees,
      }
      if (montantNum !== undefined) {
        payload.field_8 = montantNum
      } else if (montant.trim() === '') {
        // Effacement explicite du montant si le user vide le champ
        payload.field_8 = null
      }

      const result = await DCPO_LISTE_ANORMALIEService.update(String(ticket.ID), payload as never)
      if (!result.success) {
        setError('Échec de la sauvegarde : SharePoint a rejeté la mise à jour.')
        return
      }
      await onSaved()
    } catch (err) {
      console.error('Erreur sauvegarde édition manager', err)
      setError(err instanceof Error ? err.message : 'Échec de la sauvegarde.')
    } finally {
      setSaving(false)
    }
  }

  return (
    <ModalOverlay onClose={saving ? () => {} : onClose}>
      <div
        className="modal bulletin-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="bulletin-edit-title"
        style={{ maxWidth: 900 }}
      >
        <header className="bulletin-modal-header">
          <div className="bulletin-modal-title-block">
            <span className="bulletin-modal-num">{bulletin.numero}</span>
            <h2 id="bulletin-edit-title">Modifier l'anomalie clôturée</h2>
          </div>
          <button
            className="modal-close"
            onClick={onClose}
            disabled={saving}
            aria-label="Fermer sans enregistrer"
          >
            &times;
          </button>
        </header>

        <form className="bulletin-modal-body" onSubmit={handleSubmit}>
          {error && (
            <div className="bulletin-error" role="alert" style={{ marginBottom: 12 }}>
              {error}
            </div>
          )}

          {/* ─── Identification ───────────────────────────────────────── */}
          <section className="bulletin-section">
            <h3>Identification</h3>
            <div className="form-field">
              <label htmlFor="edit-title">Titre</label>
              <input
                id="edit-title"
                type="text"
                value={title}
                onChange={e => setTitle(e.target.value)}
                disabled={saving}
              />
            </div>
            {/* Chaque UserPicker sur toute la largeur : le composant affiche
                nom + email + bouton "Changer" en flex — dans un grid à 2
                colonnes, cette combinaison dépasse et le bouton se retrouve
                coupé. Full width évite tout débordement quel que soit le
                nom / email affiché. */}
            <div className="form-field">
              <label>Personne affectée</label>
              <UserPicker
                selectedName={affecteName}
                selectedEmail={affecteEmail}
                onSelect={(name, email) => { setAffecteName(name); setAffecteEmail(email) }}
                onClear={() => { setAffecteName(''); setAffecteEmail('') }}
                disabled={saving}
              />
            </div>
            <div className="form-field">
              <label>Auteur</label>
              <UserPicker
                selectedName={auteurName}
                selectedEmail={auteurEmail}
                onSelect={(name, email) => { setAuteurName(name); setAuteurEmail(email) }}
                onClear={() => { setAuteurName(''); setAuteurEmail('') }}
                disabled={saving}
              />
            </div>
          </section>

          {/* ─── Localisation ───────────────────────────────────────── */}
          <section className="bulletin-section">
            <h3>Localisation</h3>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="edit-agence">Agence</label>
                <select
                  id="edit-agence"
                  value={agence}
                  onChange={e => setAgence(e.target.value)}
                  disabled={saving}
                >
                  <option value="">— Aucune —</option>
                  {agences.map(a => (
                    <option key={a.ID} value={String(a.ID)}>{a.Title}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="edit-reseau">Réseau</label>
                <select
                  id="edit-reseau"
                  value={reseau}
                  onChange={e => setReseau(e.target.value)}
                  disabled={saving}
                >
                  <option value="">— Aucun —</option>
                  {reseaux.map(r => (
                    <option key={r.ID} value={String(r.ID)}>{r.field_1 ?? r.Title}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="edit-domaine">Domaine d'activité</label>
                <select
                  id="edit-domaine"
                  value={domaineActivite}
                  onChange={e => setDomaineActivite(e.target.value)}
                  disabled={saving}
                >
                  <option value="">— Aucun —</option>
                  {DOMAINE_ACTIVITE_OPTIONS.map(d => (
                    <option key={d} value={d}>{d}</option>
                  ))}
                </select>
              </div>
            </div>
          </section>

          {/* ─── Caractérisation ───────────────────────────────────────── */}
          <section className="bulletin-section">
            <h3>Caractérisation</h3>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="edit-classif">Classification</label>
                <select
                  id="edit-classif"
                  value={classification}
                  onChange={e => setClassification(e.target.value)}
                  disabled={saving}
                >
                  <option value="">— Aucune —</option>
                  {CLASSIFICATION_OPTIONS_EDIT.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="edit-crit">Criticité</label>
                <select
                  id="edit-crit"
                  value={criticite}
                  onChange={e => setCriticite(e.target.value)}
                  disabled={saving}
                >
                  <option value="">— Aucune —</option>
                  {CRITICITE_OPTIONS_EDIT.map(c => (
                    <option key={c} value={c}>{c}</option>
                  ))}
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="edit-nature">Type de risque</label>
                <input
                  id="edit-nature"
                  type="text"
                  value={natureRisque}
                  onChange={e => setNatureRisque(e.target.value)}
                  disabled={saving}
                />
              </div>
              <div className="form-field">
                <label htmlFor="edit-montant">Montant</label>
                <input
                  id="edit-montant"
                  type="text"
                  inputMode="decimal"
                  placeholder="0"
                  value={montant}
                  onChange={e => setMontant(e.target.value)}
                  disabled={saving}
                />
              </div>
              <div className="form-field">
                <label htmlFor="edit-type-sanction">Mode de traitement</label>
                <select
                  id="edit-type-sanction"
                  value={typeSanction}
                  onChange={e => setTypeSanction(e.target.value)}
                  disabled={saving}
                >
                  <option value="">— Aucun —</option>
                  {TYPE_SANCTION_OPTIONS.map(t => (
                    <option key={t} value={t}>{t}</option>
                  ))}
                </select>
              </div>
            </div>
          </section>

          {/* ─── Statut & dates ───────────────────────────────────────── */}
          <section className="bulletin-section">
            <h3>Statut &amp; dates</h3>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="edit-statut">Statut</label>
                <select
                  id="edit-statut"
                  value={statut}
                  onChange={e => setStatut(e.target.value)}
                  disabled={saving}
                >
                  <option value="Resolu">Resolu</option>
                  <option value="Clos">Clos</option>
                </select>
              </div>
              <div className="form-field">
                <label htmlFor="edit-date-surv">Date de survenance</label>
                <input
                  id="edit-date-surv"
                  type="date"
                  value={dateSurvenance}
                  onChange={e => setDateSurvenance(e.target.value)}
                  disabled={saving}
                />
              </div>
              <div className="form-field">
                <label htmlFor="edit-date-ouv">Date d'ouverture du ticket</label>
                <input
                  id="edit-date-ouv"
                  type="date"
                  value={dateOuverture}
                  onChange={e => setDateOuverture(e.target.value)}
                  disabled={saving}
                />
              </div>
              <div className="form-field">
                <label htmlFor="edit-date-regul">Date de régularisation</label>
                <input
                  id="edit-date-regul"
                  type="date"
                  value={dateRegul}
                  onChange={e => setDateRegul(e.target.value)}
                  disabled={saving}
                />
              </div>
              <div className="form-field">
                <label htmlFor="edit-date-cloture">Date de clôture du ticket</label>
                <input
                  id="edit-date-cloture"
                  type="date"
                  value={dateCloture}
                  min={dateOuverture || undefined}
                  onChange={e => setDateCloture(e.target.value)}
                  disabled={saving}
                />
              </div>
            </div>
          </section>

          {/* ─── Description & causes ───────────────────────────────────── */}
          <section className="bulletin-section">
            <h3>Description &amp; causes</h3>
            <div className="form-field">
              <label htmlFor="edit-description">Description consolidée</label>
              <textarea
                id="edit-description"
                rows={5}
                value={description}
                onChange={e => setDescription(e.target.value)}
                disabled={saving}
              />
              <small className="field-hint">
                Texte libre affiché en tête du bulletin. Les sections ci-dessous (cause immédiate,
                cause racine, observations, type de risque) sont ajoutées automatiquement à la
                suite lors de l'enregistrement.
              </small>
            </div>
            <div className="form-grid">
              <div className="form-field">
                <label htmlFor="edit-cause-imm">Cause immédiate</label>
                <textarea
                  id="edit-cause-imm"
                  rows={3}
                  value={causeImmediate}
                  onChange={e => setCauseImmediate(e.target.value)}
                  disabled={saving}
                />
              </div>
              <div className="form-field">
                <label htmlFor="edit-cause-rac">Cause racine</label>
                <textarea
                  id="edit-cause-rac"
                  rows={3}
                  value={causeRacine}
                  onChange={e => setCauseRacine(e.target.value)}
                  disabled={saving}
                />
              </div>
            </div>
            <div className="form-field">
              <label htmlFor="edit-actions">Actions menées</label>
              <textarea
                id="edit-actions"
                rows={3}
                value={actionsMenees}
                onChange={e => setActionsMenees(e.target.value)}
                disabled={saving}
              />
            </div>
            <div className="form-field">
              <label htmlFor="edit-obs">Observations</label>
              <textarea
                id="edit-obs"
                rows={3}
                value={observations}
                onChange={e => setObservations(e.target.value)}
                disabled={saving}
              />
            </div>
          </section>

          <footer
            className="bulletin-modal-actions no-print"
            style={{ justifyContent: 'flex-end', gap: 8, padding: '12px 0 0' }}
          >
            <button
              type="button"
              className="btn-cta btn-cta-ghost"
              onClick={onClose}
              disabled={saving}
            >
              Annuler
            </button>
            <button
              type="submit"
              className="btn-cta btn-cta-primary"
              disabled={saving}
            >
              {saving ? 'Enregistrement…' : 'Enregistrer'}
            </button>
          </footer>
        </form>
      </div>
    </ModalOverlay>
  )
}
