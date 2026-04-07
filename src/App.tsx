import { useState, useEffect } from 'react'
import { Office365UsersService } from './generated/services/Office365UsersService'
import type { GraphUser_V1 } from './generated/models/Office365UsersModel'
import Dashboard from './pages/Dashboard'
import './App.css'

function App() {
  const [page, setPage] = useState<'home' | 'dashboard'>('home')
  const [user, setUser] = useState<GraphUser_V1 | null>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    const fetchProfile = async () => {
      try {
        const result = await Office365UsersService.MyProfile_V2()
        if (result.data) {
          setUser(result.data)
        }
      } catch (err) {
        setError('Impossible de charger le profil utilisateur.')
        console.error(err)
      } finally {
        setLoading(false)
      }
    }
    fetchProfile()
  }, [])

  if (page === 'dashboard') {
    return <Dashboard onBack={() => setPage('home')} />
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
            </tbody>
          </table>
        </div>
        <button className="btn-dashboard" onClick={() => setPage('dashboard')}>
          Voir le dashboard
        </button>
      </section>
    </>
  )
}

export default App
