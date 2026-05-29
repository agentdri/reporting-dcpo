/**
 * ============================================================================
 * COMPOSANT RACINE DE L'APPLICATION — AUTHENTIFICATION & ROUTAGE TOP-LEVEL
 * ============================================================================
 *
 * Rôle :
 *   1. Récupérer le profil de l'utilisateur connecté via Office 365
 *   2. Vérifier si cet utilisateur figure dans la liste DCPO_LISTE_USER
 *   3. Vérifier si son rôle (fonction) fait partie des rôles autorisés
 *   4. Selon le résultat, afficher :
 *      - Loading : pendant la récupération
 *      - Page d'accueil avec carte profil : si non autorisé OU pas encore de
 *        clic sur "Voir le dashboard"
 *      - Dashboard : si autorisé et clic sur le bouton
 *      - Erreur : si profil introuvable ou erreur réseau
 *
 * Sécurité :
 *   La vérification se fait CÔTÉ CLIENT — c'est un garde-fou UX, pas une
 *   barrière sécurité réelle. Les permissions effectives sont gérées par
 *   SharePoint au niveau des listes (les services renvoient des erreurs si
 *   l'utilisateur n'a pas accès). Ne JAMAIS se fier uniquement à
 *   `authorized` pour cacher des données sensibles.
 * ============================================================================
 */

import { useState, useEffect } from 'react'
import { Office365UsersService } from './generated/services/Office365UsersService'
import { DCPO_LISTE_USERService } from './generated/services/DCPO_LISTE_USERService'
import type { GraphUser_V1 } from './generated/models/Office365UsersModel'
import type { DCPO_LISTE_USERRead } from './generated/models/DCPO_LISTE_USERModel'
import Dashboard from './pages/Dashboard'
import './App.css'

/**
 * Liste fermée des rôles (champ `fonction` sur DCPO_LISTE_USER) qui ont droit
 * d'accéder au dashboard. Tout autre rôle ou rôle absent → page d'accueil
 * sans bouton dashboard.
 *
 * Pour accorder l'accès à un nouveau rôle : ajouter ici ET (optionnellement)
 * adapter la logique manager dans Dashboard.tsx (constante MANAGER_ROLES).
 */
const ALLOWED_ROLES = ['Chef_Departement', 'Directeur', 'Controleur']

function App() {
  // ─── États ─────────────────────────────────────────────────────────────

  /**
   * Page actuellement affichée :
   *   - 'loading' : tant que la vérification d'identité n'est pas terminée
   *   - 'home'    : page d'accueil avec carte profil (autorisé ou non)
   *   - 'dashboard' : interface principale (réservée aux ALLOWED_ROLES)
   */
  const [page, setPage] = useState<'home' | 'dashboard' | 'loading'>('loading')

  /** Profil Microsoft Graph de l'utilisateur connecté (Office 365) */
  const [user, setUser] = useState<GraphUser_V1 | null>(null)

  /**
   * Rôle métier RÉEL (issu de DCPO_LISTE_USER) — null si user inconnu.
   * C'est la "vérité" : il sert à l'autorisation et borne les rôles que
   * l'utilisateur peut simuler (cf. activeRole).
   */
  const [userRole, setUserRole] = useState<string | null>(null)

  /**
   * Rôle EFFECTIF utilisé par l'interface (nav, permissions du dashboard).
   *
   * Par défaut égal au rôle réel. Mais un Directeur ou un Chef_Departement
   * peut le "rétrograder" temporairement via le sélecteur de la topbar pour
   * voir la plateforme comme un rôle inférieur :
   *   - Directeur        → peut simuler Chef_Departement ou Controleur
   *   - Chef_Departement → peut simuler Controleur
   *
   * C'est une simulation côté client UNIQUEMENT (aide à la vérification UX,
   * démonstration, support). Le rôle réel reste inchangé en base et
   * l'autorisation d'accès continue de s'appuyer dessus.
   */
  const [activeRole, setActiveRole] = useState<string | null>(null)

  /** True ssi user existe dans DCPO_LISTE_USER ET son rôle ∈ ALLOWED_ROLES */
  const [authorized, setAuthorized] = useState(false)

  /** True pendant la phase de récupération (init) */
  const [loading, setLoading] = useState(true)

  /** Message d'erreur à afficher si l'init échoue */
  const [error, setError] = useState<string | null>(null)

  // ─── Effet d'initialisation ───────────────────────────────────────────
  // Au montage, on récupère en série :
  //   1. Le profil Office 365 du user courant (Graph API)
  //   2. La liste complète DCPO_LISTE_USER pour matcher l'email
  //   3. Si match trouvé et rôle valide → autorisation accordée
  useEffect(() => {
    const init = async () => {
      try {
        // 1. Profil Office 365 — fournit displayName, mail, jobTitle, etc.
        const profileResult = await Office365UsersService.MyProfile_V2()
        if (!profileResult.data) {
          setError('Impossible de charger le profil utilisateur.')
          return
        }
        setUser(profileResult.data)
        // Normalisation lowercase pour comparaison case-insensitive
        const userEmail = profileResult.data.mail?.toLowerCase()

        // 2. Liste DCPO_LISTE_USER — cherche l'utilisateur par email
        const usersResult = await DCPO_LISTE_USERService.getAll()
        if (usersResult.data) {
          const match = usersResult.data.find(
            (u: DCPO_LISTE_USERRead) => u.Email?.toLowerCase() === userEmail
          )
          // 3. Vérification du rôle (fonction) — doit être dans ALLOWED_ROLES
          if (match?.fonction?.Value && ALLOWED_ROLES.includes(match.fonction.Value)) {
            setAuthorized(true)
            setUserRole(match.fonction.Value)
            // Le rôle effectif démarre toujours égal au rôle réel.
            setActiveRole(match.fonction.Value)
          }
        }
      } catch (err) {
        // Erreurs typiques : SDK non initialisé, connecteur SharePoint/O365 KO,
        // utilisateur sans licence O365, etc.
        setError('Impossible de charger le profil utilisateur.')
        console.error(err)
      } finally {
        // Quoiqu'il arrive, on sort du loading pour afficher home ou erreur
        setLoading(false)
      }
    }
    init()
  }, [])

  // ─── Effet de routage ─────────────────────────────────────────────────
  // Quand le loading se termine, on décide automatiquement :
  //   - Si autorisé → on entre directement dans le dashboard
  //   - Sinon → on reste sur 'home'
  useEffect(() => {
    if (!loading) {
      setPage(authorized ? 'dashboard' : 'home')
    }
  }, [loading, authorized])

  // ─── Rendu : page Dashboard ───────────────────────────────────────────
  if (page === 'dashboard') {
    // Garde de sécurité : si on a basculé sur dashboard mais que l'autorisation
    // a changé entre-temps (très rare mais possible), on revient sur 'home'.
    if (!authorized) {
      setPage('home')
      return null
    }
    return (
      <Dashboard
        userName={user?.displayName}
        // Le dashboard pilote sa nav/permissions sur le rôle EFFECTIF
        userRole={activeRole ?? userRole ?? undefined}
        // Le rôle RÉEL borne les rôles simulables (sélecteur topbar)
        realUserRole={userRole ?? undefined}
        onRoleChange={setActiveRole}
        userEmail={user?.mail ?? undefined}
      />
    )
  }

  // ─── Rendu : écran de chargement ──────────────────────────────────────
  if (loading) {
    return (
      <section id="center">
        <p>Chargement du profil...</p>
      </section>
    )
  }

  // ─── Rendu : écran d'erreur (profil introuvable / SDK KO) ─────────────
  if (error || !user) {
    return (
      <section id="center">
        <p>{error ?? 'Utilisateur introuvable.'}</p>
      </section>
    )
  }

  // ─── Rendu : page d'accueil (carte profil + bouton dashboard) ─────────
  // Affichée pour tout user connecté avec O365, qu'il soit autorisé ou non.
  // Les non-autorisés voient le message "Vous n'avez pas les droits..."
  // Les autorisés voient le bouton "Voir le dashboard" pour entrer dans l'app.
  return (
    <>
      <section id="center">
        <h1 className="welcome-title">Bienvenue, {user.displayName}</h1>
        <div className="user-card">
          {/* Tableau récapitulatif des infos utilisateur (Graph) */}
          <table className="user-info">
            <tbody>
              <tr><td className="label">Email</td><td>{user.mail ?? '-'}</td></tr>
              <tr><td className="label">Poste</td><td>{user.jobTitle ?? '-'}</td></tr>
              <tr><td className="label">Departement</td><td>{user.department ?? '-'}</td></tr>
              <tr>
                <td className="label">Telephone</td>
                {/* Fallback en cascade : mobile → premier business phone → '-' */}
                <td>{user.mobilePhone ?? user.businessPhones?.[0] ?? '-'}</td>
              </tr>
              <tr><td className="label">Bureau</td><td>{user.officeLocation ?? '-'}</td></tr>
              <tr><td className="label">Ville</td><td>{user.city ?? '-'}</td></tr>
              <tr><td className="label">Pays</td><td>{user.country ?? '-'}</td></tr>
              {/* Le rôle métier n'est affiché que si on l'a trouvé en BDD */}
              {userRole && (
                <tr>
                  <td className="label">Role</td>
                  <td>{userRole}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {/* Action conditionnelle : bouton si autorisé, message d'info sinon */}
        {authorized ? (
          <button className="btn-dashboard" onClick={() => setPage('dashboard')}>
            Voir le dashboard
          </button>
        ) : (
          <p className="no-access">Vous n'avez pas les droits pour acceder au dashboard.</p>
        )}
      </section>
    </>
  )
}

export default App
