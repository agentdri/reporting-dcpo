/**
 * ============================================================================
 * POINT D'ENTRÉE DE L'APPLICATION REACT
 * ============================================================================
 *
 * Ce fichier est la racine du bundle Vite. Il est référencé par index.html
 * via <script type="module" src="/src/main.tsx">. C'est le tout premier
 * code React exécuté quand l'app charge dans le navigateur.
 *
 * Responsabilités :
 *   1. Importer les styles globaux (./index.css)
 *   2. Initialiser le SDK Power Apps (via PowerProvider)
 *   3. Monter le composant racine <App /> sur l'élément DOM #root
 *   4. Activer StrictMode (vérifications React supplémentaires en dev)
 * ============================================================================
 */

import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './theme.css'
import './index.css'
import App from './App.tsx'
import PowerProvider from './PowerProvider.tsx'

// createRoot = API React 18+ pour le rendu concurrent
// document.getElementById('root')! → le `!` indique à TypeScript que l'élément
// existe forcément (défini dans index.html). Si absent, l'app crash au démarrage.
createRoot(document.getElementById('root')!).render(
  // StrictMode :
  //   - Double l'invocation des effets/render en DEV pour détecter les
  //     side-effects accidentels
  //   - Aucun impact en PRODUCTION (Vite l'élimine au build)
  <StrictMode>
    {/* PowerProvider : enveloppe l'app pour initialiser le SDK Power Platform
        AVANT que les pages tentent d'appeler les services SharePoint */}
    <PowerProvider>
      <App />
    </PowerProvider>
  </StrictMode>,
)
