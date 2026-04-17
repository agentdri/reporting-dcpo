import { useState, useEffect } from 'react'
import { Office365UsersService } from './generated/services/Office365UsersService'
import { DCPO_LISTE_USERService } from './generated/services/DCPO_LISTE_USERService'
import type { GraphUser_V1 } from './generated/models/Office365UsersModel'
import type { DCPO_LISTE_USERRead } from './generated/models/DCPO_LISTE_USERModel'
import Dashboard from './pages/Dashboard'
import './App.css'

const ALLOWED_ROLES = ['Chef_Departement', 'Directeur', 'Controleur']

function App() {
  const [page, setPage] = useState<'home' | 'dashboard' | 'loading'>('loading')
  const [user, setUser] = useState<GraphUser_V1 | null>(null)
  const [userRole, setUserRole] = useState<string | null>(null)
  const [authorized, setAuthorized] = useState(false)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    const init = async () => {
      try {
        const profileResult = await Office365UsersService.MyProfile_V2()
        if (!profileResult.data) {
          setError('Impossible de charger le profil utilisateur.')
          return
        }
        setUser(profileResult.data)
        const userEmail = profileResult.data.mail?.toLowerCase()

        const usersResult = await DCPO_LISTE_USERService.getAll()
        if (usersResult.data) {
          const match = usersResult.data.find(
            (u: DCPO_LISTE_USERRead) => u.Email?.toLowerCase() === userEmail
          )
          if (match?.fonction?.Value && ALLOWED_ROLES.includes(match.fonction.Value)) {
            setAuthorized(true)
            setUserRole(match.fonction.Value)
          }
        }
      } catch (err) {
        setError('Impossible de charger le profil utilisateur.')
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    init()
  }, [])

  useEffect(() => {
    if (!loading) {
      setPage(authorized ? 'dashboard' : 'home')
    }
  }, [loading, authorized])

  if (page === 'dashboard') {
    if (!authorized) {
      setPage('home')
      return null
    }
    return <Dashboard userName={user?.displayName} userRole={userRole ?? undefined} userEmail={user?.mail ?? undefined} />
  }

  if (loading) {
    return (
      <section id="center">
        <p>Chargement du profil...</p>
      </section>
    )
  }

  if (error || !user) {
    return (
      <section id="center">
        <p>{error ?? 'Utilisateur introuvable.'}</p>
      </section>
    )
  }

  return (
    <>
      <section id="center">
        <h1 className="welcome-title">Bienvenue, {user.displayName}</h1>
        <div className="user-card">
          <table className="user-info">
            <tbody>
              <tr>
                <td className="label">Email</td>
                <td>{user.mail ?? '-'}</td>
              </tr>
              <tr>
                <td className="label">Poste</td>
                <td>{user.jobTitle ?? '-'}</td>
              </tr>
              <tr>
                <td className="label">Departement</td>
                <td>{user.department ?? '-'}</td>
              </tr>
              <tr>
                <td className="label">Telephone</td>
                <td>{user.mobilePhone ?? user.businessPhones?.[0] ?? '-'}</td>
              </tr>
              <tr>
                <td className="label">Bureau</td>
                <td>{user.officeLocation ?? '-'}</td>
              </tr>
              <tr>
                <td className="label">Ville</td>
                <td>{user.city ?? '-'}</td>
              </tr>
              <tr>
                <td className="label">Pays</td>
                <td>{user.country ?? '-'}</td>
              </tr>
              {userRole && (
                <tr>
                  <td className="label">Role</td>
                  <td>{userRole}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
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