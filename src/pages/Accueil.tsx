/**
 * ============================================================================
 * Accueil — Vue d'ensemble (tableau de bord d'entrée)
 * ============================================================================
 *
 * Page d'accueil de l'application, dans la charte « modernisée » : une rangée de
 * cartes-chiffres (les indicateurs principaux des anomalies) surmontée d'un mot
 * de bienvenue, puis une grille de graphiques Recharts qui « traduisent » ces
 * chiffres :
 *
 *   - Donut    : répartition des anomalies par statut (Ouvert / En cours /
 *                Résolu / Clos), total au centre.
 *   - Barres   : nombre d'anomalies par niveau de criticité.
 *   - Aire     : anomalies déclarées par mois sur les 12 derniers mois (tendance).
 *   - Barres ↔ : top des classifications les plus fréquentes.
 *
 * Données : toutes les anomalies (tous statuts) via DCPO_LISTE_ANORMALIEService.
 * Contrairement à l'onglet « Anomalies en cours » qui filtre côté serveur sur
 * les statuts ouverts, l'accueil agrège l'ensemble pour donner une vision
 * globale.
 */
import { useEffect, useMemo, useRef, useState } from 'react'
import {
  ResponsiveContainer,
  PieChart,
  Pie,
  Cell,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  AreaChart,
  Area,
} from 'recharts'
import { DCPO_LISTE_ANORMALIEService } from '../generated/services/DCPO_LISTE_ANORMALIEService'
import type { DCPO_LISTE_ANORMALIERead } from '../generated/models/DCPO_LISTE_ANORMALIEModel'
import { getAllPages } from '../lib/sharePointPaging'
import { formatMontantCompact } from '../lib/formatters'
import { useFadeUp, useStagger } from '../lib/useGsap'
import './Accueil.css'

/** Palette alignée sur les accents de la charte (cf. theme.css / stat-cards). */
const C = {
  dark: '#141c2e',
  red: '#d42235',
  orange: '#e08a1e',
  blue: '#3465d6',
  green: '#1b8a4f',
  purple: '#8255c9',
  grid: '#e5e9f2',
  muted: '#8891b0',
}

/** Style commun des info-bulles Recharts (carte blanche arrondie, ombre douce). */
const TOOLTIP_STYLE = {
  borderRadius: 10,
  border: '1px solid #e5e9f2',
  fontSize: 12,
  boxShadow: '0 6px 18px rgba(20,28,46,.10)',
  padding: '8px 12px',
} as const

interface AccueilProps {
  userName?: string
  userRole?: string
  userEmail?: string
  /** Navigue vers un autre onglet du Dashboard (cartes cliquables). */
  onNavigate?: (tab: string) => void
}

/** Salutation selon l'heure locale. */
function greeting(): string {
  const h = new Date().getHours()
  if (h < 12) return 'Bonjour'
  if (h < 18) return 'Bon après-midi'
  return 'Bonsoir'
}

export default function Accueil({ userName, onNavigate }: AccueilProps) {
  const [items, setItems] = useState<DCPO_LISTE_ANORMALIERead[]>([])
  const [loading, setLoading] = useState(true)

  // ─── Chargement de TOUTES les anomalies (tous statuts) ─────────────────
  useEffect(() => {
    let alive = true
    ;(async () => {
      try {
        const rows = await getAllPages<DCPO_LISTE_ANORMALIERead>(
          DCPO_LISTE_ANORMALIEService,
          { orderBy: ['Created desc'] },
        )
        if (alive) setItems(rows)
      } catch (err) {
        console.error('Accueil : échec du chargement des anomalies', err)
      } finally {
        if (alive) setLoading(false)
      }
    })()
    return () => {
      alive = false
    }
  }, [])

  // ─── Indicateurs principaux (cartes) ───────────────────────────────────
  const stats = useMemo(() => {
    const countBy = (s: string) => items.filter(i => i.field_10 === s).length
    return {
      total: items.length,
      ouvert: countBy('Ouvert'),
      enCours: countBy('En cours'),
      resolu: countBy('Resolu'),
      clos: countBy('Clos'),
      montant: items.reduce((sum, i) => sum + (i.field_8 ?? 0), 0),
    }
  }, [items])

  // ─── Donut : répartition par statut ────────────────────────────────────
  const statusData = useMemo(
    () => [
      { name: 'Ouvert', value: stats.ouvert, color: C.red },
      { name: 'En cours', value: stats.enCours, color: C.orange },
      { name: 'Résolu', value: stats.resolu, color: C.blue },
      { name: 'Clos', value: stats.clos, color: C.green },
    ],
    [stats],
  )
  const statusSlices = statusData.filter(d => d.value > 0)

  // ─── Barres : anomalies par criticité ──────────────────────────────────
  const critData = useMemo(() => {
    const defs = [
      { name: 'Faible', color: C.green },
      { name: 'Moyenne', color: C.blue },
      { name: 'Haute', color: C.orange },
      { name: 'Critique', color: C.red },
    ]
    return defs.map(d => ({
      ...d,
      value: items.filter(
        i => (i.criticiteAnomalie ?? '').trim().toLowerCase() === d.name.toLowerCase(),
      ).length,
    }))
  }, [items])

  // ─── Aire : anomalies déclarées par mois (12 derniers mois) ─────────────
  const trendData = useMemo(() => {
    const now = new Date()
    const buckets: { key: string; label: string; value: number }[] = []
    for (let k = 11; k >= 0; k--) {
      const d = new Date(now.getFullYear(), now.getMonth() - k, 1)
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
      buckets.push({ key, label: d.toLocaleDateString('fr-FR', { month: 'short' }), value: 0 })
    }
    const idx = new Map(buckets.map((b, i) => [b.key, i]))
    for (const it of items) {
      const raw = it.field_0 || it.Created
      if (!raw) continue
      const d = new Date(raw)
      if (Number.isNaN(d.getTime())) continue
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`
      const i = idx.get(key)
      if (i != null) buckets[i].value += 1
    }
    return buckets
  }, [items])

  // ─── Barres horizontales : top classifications ─────────────────────────
  const classifData = useMemo(() => {
    const m = new Map<string, number>()
    for (const it of items) {
      const c = (it.field_5 ?? '').trim() || 'Non classé'
      m.set(c, (m.get(c) ?? 0) + 1)
    }
    return [...m.entries()]
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 6)
  }, [items])

  // ─── Animations d'entrée (GSAP) ────────────────────────────────────────
  const pageRef = useRef<HTMLDivElement>(null)
  useFadeUp(pageRef, [loading])
  useStagger(pageRef, { selector: '.stat-card', deps: [loading, items.length] })

  const hasData = items.length > 0

  return (
    <div className="accueil" ref={pageRef}>
      {/* ─── En-tête de bienvenue ──────────────────────────────────────── */}
      <div className="accueil-hero">
        <div>
          <div className="accueil-hello">
            {greeting()}{userName ? `, ${userName.split(' ')[0]}` : ''} 👋
          </div>
          <div className="accueil-sub">
            Vue d'ensemble des anomalies détectées lors des contrôles
          </div>
        </div>
      </div>

      {/* ─── Cartes-chiffres (indicateurs principaux) ──────────────────── */}
      <div className="stats-cards">
        <button type="button" className="stat-card total" onClick={() => onNavigate?.('anomalie')}>
          <span className="stat-value">{loading ? '—' : stats.total}</span>
          <span className="stat-label">Total</span>
        </button>
        <button type="button" className="stat-card ouvert" onClick={() => onNavigate?.('anomalie')}>
          <span className="stat-value">{loading ? '—' : stats.ouvert}</span>
          <span className="stat-label">Ouvert</span>
        </button>
        <button type="button" className="stat-card en-cours" onClick={() => onNavigate?.('anomalie')}>
          <span className="stat-value">{loading ? '—' : stats.enCours}</span>
          <span className="stat-label">En cours</span>
        </button>
        <button type="button" className="stat-card resolu" onClick={() => onNavigate?.('bulletins')}>
          <span className="stat-value">{loading ? '—' : stats.resolu}</span>
          <span className="stat-label">Résolu</span>
        </button>
        <button type="button" className="stat-card clos" onClick={() => onNavigate?.('bulletins')}>
          <span className="stat-value">{loading ? '—' : stats.clos}</span>
          <span className="stat-label">Clos</span>
        </button>
        <div className="stat-card montant">
          <span className="stat-value">{loading ? '—' : formatMontantCompact(stats.montant)}</span>
          <span className="stat-label">Montant total</span>
        </div>
      </div>

      {/* ─── Grille de graphiques ──────────────────────────────────────── */}
      <div className="accueil-charts">
        {/* Donut statut */}
        <section className="chart-panel">
          <header className="chart-head">
            <h3>Répartition par statut</h3>
          </header>
          {hasData && statusSlices.length > 0 ? (
            <div className="donut-wrap">
              <ResponsiveContainer width="100%" height={220}>
                <PieChart>
                  <Pie
                    data={statusSlices}
                    dataKey="value"
                    nameKey="name"
                    innerRadius={62}
                    outerRadius={92}
                    paddingAngle={2}
                    stroke="none"
                  >
                    {statusSlices.map(d => (
                      <Cell key={d.name} fill={d.color} />
                    ))}
                  </Pie>
                  <Tooltip contentStyle={TOOLTIP_STYLE} />
                </PieChart>
              </ResponsiveContainer>
              <div className="donut-center">
                <span className="donut-total">{stats.total}</span>
                <span className="donut-total-label">anomalies</span>
              </div>
            </div>
          ) : (
            <Empty loading={loading} />
          )}
          {/* Légende avec compteurs */}
          <ul className="chart-legend">
            {statusData.map(d => (
              <li key={d.name}>
                <span className="legend-dot" style={{ background: d.color }} />
                <span className="legend-name">{d.name}</span>
                <span className="legend-val">{d.value}</span>
              </li>
            ))}
          </ul>
        </section>

        {/* Barres criticité */}
        <section className="chart-panel">
          <header className="chart-head">
            <h3>Anomalies par criticité</h3>
          </header>
          {hasData ? (
            <ResponsiveContainer width="100%" height={260}>
              <BarChart data={critData} margin={{ top: 8, right: 8, left: -16, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke={C.grid} vertical={false} />
                <XAxis dataKey="name" tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <YAxis allowDecimals={false} tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <Tooltip cursor={{ fill: 'rgba(20,28,46,.04)' }} contentStyle={TOOLTIP_STYLE} />
                <Bar dataKey="value" name="Anomalies" radius={[6, 6, 0, 0]} maxBarSize={56}>
                  {critData.map(d => (
                    <Cell key={d.name} fill={d.color} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <Empty loading={loading} />
          )}
        </section>

        {/* Aire tendance (pleine largeur) */}
        <section className="chart-panel span-2">
          <header className="chart-head">
            <h3>Anomalies déclarées — 12 derniers mois</h3>
          </header>
          {hasData ? (
            <ResponsiveContainer width="100%" height={240}>
              <AreaChart data={trendData} margin={{ top: 8, right: 12, left: -16, bottom: 0 }}>
                <defs>
                  <linearGradient id="trendFill" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={C.red} stopOpacity={0.28} />
                    <stop offset="100%" stopColor={C.red} stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke={C.grid} vertical={false} />
                <XAxis dataKey="label" tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <YAxis allowDecimals={false} tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={TOOLTIP_STYLE} />
                <Area
                  type="monotone"
                  dataKey="value"
                  name="Anomalies"
                  stroke={C.red}
                  strokeWidth={2.5}
                  fill="url(#trendFill)"
                  dot={{ r: 3, fill: C.red, strokeWidth: 0 }}
                  activeDot={{ r: 5 }}
                />
              </AreaChart>
            </ResponsiveContainer>
          ) : (
            <Empty loading={loading} />
          )}
        </section>

        {/* Top classifications (pleine largeur) */}
        <section className="chart-panel span-2">
          <header className="chart-head">
            <h3>Top classifications</h3>
          </header>
          {hasData && classifData.length > 0 ? (
            <ResponsiveContainer width="100%" height={Math.max(160, classifData.length * 42)}>
              <BarChart
                data={classifData}
                layout="vertical"
                margin={{ top: 4, right: 24, left: 8, bottom: 4 }}
              >
                <CartesianGrid strokeDasharray="3 3" stroke={C.grid} horizontal={false} />
                <XAxis type="number" allowDecimals={false} tick={{ fontSize: 12, fill: C.muted }} axisLine={false} tickLine={false} />
                <YAxis
                  type="category"
                  dataKey="name"
                  width={150}
                  tick={{ fontSize: 12, fill: C.dark }}
                  axisLine={false}
                  tickLine={false}
                />
                <Tooltip cursor={{ fill: 'rgba(20,28,46,.04)' }} contentStyle={TOOLTIP_STYLE} />
                <Bar dataKey="value" name="Anomalies" fill={C.blue} radius={[0, 6, 6, 0]} maxBarSize={26} />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <Empty loading={loading} />
          )}
        </section>
      </div>
    </div>
  )
}

/** Encart d'état vide / de chargement, partagé par tous les graphiques. */
function Empty({ loading }: { loading: boolean }) {
  return (
    <div className="chart-empty">
      {loading ? 'Chargement…' : 'Aucune donnée à afficher.'}
    </div>
  )
}
