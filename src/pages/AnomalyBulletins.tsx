import { useEffect, useMemo, useState } from 'react'
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
import './AnomalyBulletins.css'

const STATUS_OPTIONS: BulletinStatus[] = ['Tous', 'Resolu', 'Clos']
const CLASSIFICATION_OPTIONS = ['Operationnel', 'Fraude', 'Commercial']
const CRITICITE_OPTIONS = ['Faible', 'Moyenne', 'Haute', 'Critique']

export default function AnomalyBulletins() {
  const [tickets, setTickets] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [agences, setAgences] = useState<DCPO_LISTE_AGENCESRead[]>([])
  const [reseaux, setReseaux] = useState<DCPO_LISTE_RESEAUXRead[]>([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)
  const [filters, setFilters] = useState<BulletinFilters>(EMPTY_BULLETIN_FILTERS)
  const [selected, setSelected] = useState<ConsolidatedBulletin | null>(null)

  const fetchData = async () => {
    setLoading(true)
    setError(null)
    try {
      const [ticketsRes, agencesRows, reseauxRows] = await Promise.all([
        DCPO_LISTE_ANORMALIEService.getAll({
          orderBy: ['Created desc'],
          filter: "field_10 eq 'Resolu' or field_10 eq 'Clos'",
        }),
        loadAgences(),
        loadReseaux(),
      ])
      setAgences(agencesRows)
      setReseaux(reseauxRows)
      if (ticketsRes.data) {
        setTickets(ticketsRes.data.filter(isBulletinTicket))
      } else {
        setTickets([])
      }
    } catch (err) {
      console.error('Erreur chargement bulletins', err)
      setError('Impossible de charger les bulletins.')
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    fetchData()
  }, [])

  const bulletins = useMemo(
    () => tickets.map(t => buildConsolidatedBulletin(t, agences, reseaux)),
    [tickets, agences, reseaux],
  )

  const filtered = useMemo(() => applyBulletinFilters(bulletins, filters), [bulletins, filters])

  const stats = useMemo(() => {
    const resolus = filtered.filter(b => b.statut === 'Resolu').length
    const clos = filtered.filter(b => b.statut === 'Clos').length
    const totalMontant = filtered.reduce((s, b) => s + (b.montant ?? 0), 0)
    const delais = filtered.map(b => b.delayDays ?? 0).filter(n => n > 0)
    const avgDelay = delais.length ? Math.round(delais.reduce((s, n) => s + n, 0) / delais.length) : 0
    return { total: filtered.length, resolus, clos, totalMontant, avgDelay }
  }, [filtered])

  const updateFilter = <K extends keyof BulletinFilters>(key: K, value: BulletinFilters[K]) => {
    setFilters(prev => ({ ...prev, [key]: value }))
  }

  const resetFilters = () => setFilters(EMPTY_BULLETIN_FILTERS)

  const printBulletin = () => {
    window.print()
  }

  return (
    <>
      <div className="content-header">
        <h2>Bulletins d'anomalies</h2>
        <button className="btn-add" onClick={fetchData} disabled={loading} type="button">
          {loading ? 'Actualisation...' : 'Actualiser'}
        </button>
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
          <span className="stat-value">{stats.totalMontant.toLocaleString('fr-FR')}</span>
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
            <label>Clôture du</label>
            <input type="date" value={filters.closureFrom} onChange={e => updateFilter('closureFrom', e.target.value)} />
          </div>
          <div className="filter-field">
            <label>Clôture au</label>
            <input type="date" value={filters.closureTo} onChange={e => updateFilter('closureTo', e.target.value)} />
          </div>
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
        <div className="bulletins-grid">
          {filtered.map(b => (
            <BulletinCard key={b.ticket.ID ?? b.numero} bulletin={b} onOpen={() => setSelected(b)} />
          ))}
        </div>
      )}

      {selected && (
        <BulletinModal
          bulletin={selected}
          onClose={() => setSelected(null)}
          onPrint={printBulletin}
        />
      )}
    </>
  )
}

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

interface BulletinModalProps {
  bulletin: ConsolidatedBulletin
  onClose: () => void
  onPrint: () => void
}

function BulletinModal({ bulletin, onClose, onPrint }: BulletinModalProps) {
  const attachments = getTicketAttachments(bulletin.ticket)

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose()
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onClose])

  return (
    <div className="modal-overlay bulletin-modal-overlay" onClick={onClose}>
      <div
        className="modal bulletin-modal"
        onClick={e => e.stopPropagation()}
        role="dialog"
        aria-modal="true"
        aria-labelledby="bulletin-modal-title"
      >
        <header className="bulletin-modal-header">
          <div>
            <span className="bulletin-modal-num">{bulletin.numero}</span>
            <h2 id="bulletin-modal-title">{bulletin.titre}</h2>
          </div>
          <div className="bulletin-modal-actions no-print">
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
    </div>
  )
}
