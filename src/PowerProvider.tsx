/**
 * ============================================================================
 * POWER PROVIDER — INITIALISATION DU SDK POWER PLATFORM
 * ============================================================================
 *
 * Composant wrapper qui doit englober l'arbre React de l'app.
 *
 * Rôle :
 *   - Appelle `initialize()` du SDK @pa-client/power-code-sdk au montage
 *   - Cet appel établit la connexion avec l'environnement Power Platform
 *     (environment ID, app ID, connecteurs SharePoint/Office365 configurés
 *      dans power.config.json)
 *   - Sans cette initialisation, les services générés (DCPO_LISTE_*Service,
 *     Office365UsersService) lèveraient des erreurs en tentant de communiquer
 *
 * Le composant rend simplement ses enfants — il n'ajoute pas de DOM visible.
 * Il sert uniquement à exécuter l'effet d'initialisation au bon moment du
 * cycle de vie React (après le premier render, avant que les pages enfants
 * ne déclenchent leurs propres useEffect de chargement de données).
 * ============================================================================
 */

import { initialize } from "@pa-client/power-code-sdk/lib/Lifecycle";
import { useEffect, type ReactNode } from "react";

/**
 * Props du provider :
 *   - children : tout l'arbre React qui sera rendu (typiquement <App />)
 */
interface PowerProviderProps {
  children: ReactNode;
}

export default function PowerProvider({ children }: PowerProviderProps) {
  // Effect avec [] = exécuté UNE seule fois au montage du composant.
  // En StrictMode + dev, React le double-invoque ; le SDK gère cette
  // ré-entrance sans casser la connexion.
  useEffect(() => {
    const initApp = async () => {
      try {
        // initialize() :
        //   - Lit la config power (appId, environmentId, connectionReferences)
        //   - Crée les clients HTTP authentifiés vers les datasources
        //   - Doit terminer AVANT tout appel à xxxService.getAll() / .create()
        await initialize();
        console.log('Power Platform SDK initialized successfully');
      } catch (error) {
        // Erreur courante : connecteurs non autorisés, environnement indisponible.
        // On log mais on ne bloque pas l'app — c'est l'affaire de App.tsx
        // de gérer le cas où les services renvoient des erreurs au chargement.
        console.error('Failed to initialize Power Platform SDK:', error);
      }
    };

    initApp();
  }, []);

  // Pas de wrapper DOM — le provider est invisible visuellement.
  return <>{children}</>;
}
