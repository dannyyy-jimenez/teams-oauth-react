"use client";

import { createContext, useContext, useState, useEffect, ReactNode } from "react";
import { MsalProvider, useMsal, useIsAuthenticated } from "@azure/msal-react";
import { PublicClientApplication, EventType, AccountInfo, InteractionStatus } from "@azure/msal-browser";
import { Configuration } from "@azure/msal-browser";

// ============================================================================
// Types
// ============================================================================

/**
 * Configuration options for Teams authentication
 *
 * @example
 * ```tsx
 * const config: TeamsAuthConfig = {
 *   clientId: "your-azure-client-id",
 *   tenantId: "common", // or your specific tenant ID
 *   redirectUri: "http://localhost:3000",
 *   scopes: ["User.Read", "Team.ReadBasic.All"],
 *   cacheLocation: "sessionStorage"
 * };
 * ```
 */
export interface TeamsAuthConfig {
  /** Azure AD Application (client) ID - Required */
  clientId: string;
  /** Azure AD Directory (tenant) ID. Defaults to "common" for multi-tenant apps */
  tenantId?: string;
  /** OAuth redirect URI. Defaults to window.location.origin */
  redirectUri?: string;
  /** Permission scopes to request. Defaults to basic Teams read permissions */
  scopes?: string[];
  /** Where to store authentication tokens. Defaults to "sessionStorage" */
  cacheLocation?: "localStorage" | "sessionStorage";
}

/**
 * Context value provided by TeamsAuthProvider
 * Access this via the useTeamsAuth() hook
 *
 * @example
 * ```tsx
 * const { isAuthenticated, login, logout, user } = useTeamsAuth();
 * ```
 */
export interface TeamsContextValue {
  // Auth state
  /** Whether the user is currently authenticated */
  isAuthenticated: boolean;
  /** Whether an authentication operation is in progress */
  isLoading: boolean;
  /** Current user account information, null if not authenticated */
  user: AccountInfo | null;
  /** Current MSAL interaction status */
  inProgress: InteractionStatus;

  // Auth methods
  /**
   * Sign in with Microsoft account using redirect flow
   * @throws Error if login fails
   */
  login: () => Promise<void>;
  /**
   * Sign out and redirect to logout page
   * @throws Error if logout fails
   */
  logout: () => Promise<void>;

  // Token methods
  /**
   * Get an access token for the specified scopes
   * @param scopes - Optional array of permission scopes. Uses default scopes if not provided
   * @returns Access token string or null if acquisition fails
   */
  getAccessToken: (scopes?: string[]) => Promise<string | null>;

  // Graph API helper
  /**
   * Call Microsoft Graph API with automatic token handling
   * @param endpoint - Graph API endpoint (e.g., "/me" or "me/joinedTeams")
   * @param method - HTTP method. Defaults to "GET"
   * @param body - Request body for POST/PATCH requests
   * @returns API response data
   * @throws Error if API call fails
   */
  callGraphAPI: <T = unknown>(endpoint: string, method?: "GET" | "POST" | "PATCH" | "DELETE", body?: unknown) => Promise<T>;

  // Error state
  /** Current error message, null if no error */
  error: string | null;
  /** Clear the current error message */
  clearError: () => void;
}

// ============================================================================
// Context
// ============================================================================

/**
 * React context for Teams authentication
 * @internal Use useTeamsAuth() hook instead of accessing directly
 */
const TeamsContext = createContext<TeamsContextValue | undefined>(undefined);

// ============================================================================
// Default Scopes
// ============================================================================

/**
 * Default permission scopes for Teams API access
 * Includes basic read permissions for user profile, teams, channels, and messages
 */
const DEFAULT_SCOPES = [
  "User.Read",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
];

// ============================================================================
// MSAL Config Builder
// ============================================================================

/**
 * Creates MSAL configuration from Teams auth config
 * @internal
 * @param config - Teams authentication configuration
 * @returns MSAL configuration object
 */
function createMsalConfig(config: TeamsAuthConfig): Configuration {
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId || "common"}`,
      redirectUri: config.redirectUri || (typeof window !== "undefined" ? window.location.origin : "http://localhost:3000"),
      postLogoutRedirectUri: config.redirectUri || (typeof window !== "undefined" ? window.location.origin : "http://localhost:3000"),
    },
    cache: {
      cacheLocation: config.cacheLocation || "sessionStorage",
    },
  };
}

// ============================================================================
// Inner Provider (uses MSAL hooks)
// ============================================================================

/**
 * Props for the inner provider component
 * @internal
 */
interface TeamsAuthInnerProviderProps {
  children: ReactNode;
  scopes: string[];
}

/**
 * Inner provider that uses MSAL hooks and provides Teams context
 * @internal
 * @param props - Component props
 */
function TeamsAuthInnerProvider({ children, scopes }: TeamsAuthInnerProviderProps) {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const [error, setError] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);

  const user = accounts.length > 0 ? accounts[0] : null;

  // Login method
  const login = async () => {
    try {
      setError(null);
      setIsLoading(true);
      await instance.loginRedirect({
        scopes: scopes,
      });
    } catch (err) {
      console.error("Login failed:", err);
      const error = err as { message?: string };
      setError(error.message || "Login failed");
      setIsLoading(false);
    }
  };

  // Logout method
  const logout = async () => {
    try {
      setError(null);
      await instance.logoutRedirect();
    } catch (err) {
      console.error("Logout failed:", err);
      const error = err as { message?: string };
      setError(error.message || "Logout failed");
    }
  };

  // Get access token
  const getAccessToken = async (requestScopes?: string[]): Promise<string | null> => {
    if (accounts.length === 0) {
      setError("No account found. Please sign in.");
      return null;
    }

    const scopesToUse = requestScopes || scopes;

    try {
      // Try silent token acquisition
      const response = await instance.acquireTokenSilent({
        scopes: scopesToUse,
        account: accounts[0],
      });
      return response.accessToken;
    } catch {
      console.warn("⚠️ Silent token acquisition failed, using redirect...");
      try {
        await instance.acquireTokenRedirect({
          scopes: scopesToUse,
        });
        return null;
      } catch (redirectErr) {
        console.error("❌ Token acquisition failed:", redirectErr);
        const err = redirectErr as { message?: string };
        setError(err.message || "Failed to acquire token");
        return null;
      }
    }
  };

  // Call Graph API helper
  const callGraphAPI = async <T = unknown,>(
    endpoint: string,
    method: "GET" | "POST" | "PATCH" | "DELETE" = "GET",
    body?: unknown
  ): Promise<T> => {
    const token = await getAccessToken();

    if (!token) {
      throw new Error("No access token available");
    }

    const url = endpoint.startsWith("http")
      ? endpoint
      : `https://graph.microsoft.com/v1.0${endpoint.startsWith("/") ? "" : "/"}${endpoint}`;

    const options: RequestInit = {
      method,
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json",
      },
    };

    if (body && (method === "POST" || method === "PATCH")) {
      options.body = JSON.stringify(body);
    }

    const response = await fetch(url, options);

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      throw new Error(errorData.error?.message || `API call failed: ${response.statusText}`);
    }

    return response.json();
  };

  const clearError = () => setError(null);

  const contextValue: TeamsContextValue = {
    isAuthenticated,
    isLoading: isLoading || inProgress !== InteractionStatus.None,
    user,
    inProgress,
    login,
    logout,
    getAccessToken,
    callGraphAPI,
    error,
    clearError,
  };

  return (
    <TeamsContext.Provider value={contextValue}>
      {children}
    </TeamsContext.Provider>
  );
}

// ============================================================================
// Main Provider Component
// ============================================================================

/**
 * Props for TeamsAuthProvider component
 */
interface TeamsAuthProviderProps {
  /** Authentication configuration */
  config: TeamsAuthConfig;
  /** Child components to render within the provider */
  children: ReactNode;
  /** Optional custom loading component to show during initialization */
  loadingComponent?: ReactNode;
}

/**
 * Teams Authentication Provider Component
 *
 * Wraps your application and provides Microsoft Teams OAuth authentication
 * with automatic token management and Graph API integration.
 *
 * @component
 * @example
 * ```tsx
 * // In your app/layout.tsx
 * import { TeamsAuthProvider } from '@your-org/teams-oauth-plugin';
 *
 * export default function RootLayout({ children }) {
 *   return (
 *     <html>
 *       <body>
 *         <TeamsAuthProvider
 *           config={{
 *             clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID!,
 *             tenantId: "common",
 *             scopes: ["User.Read", "Team.ReadBasic.All"]
 *           }}
 *         >
 *           {children}
 *         </TeamsAuthProvider>
 *       </body>
 *     </html>
 *   );
 * }
 * ```
 *
 * @param props - Component props
 * @returns Provider component that wraps your application
 */
export function TeamsAuthProvider({ config, children, loadingComponent }: TeamsAuthProviderProps) {
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | null>(null);
  const [isInitialized, setIsInitialized] = useState(false);

  const scopes = config.scopes || DEFAULT_SCOPES;

  useEffect(() => {
    const initializeMsal = async () => {
      try {
        const msalConfig = createMsalConfig(config);
        const instance = new PublicClientApplication(msalConfig);

        await instance.initialize();

        // Set active account if exists
        const accounts = instance.getAllAccounts();
        if (accounts.length > 0) {
          instance.setActiveAccount(accounts[0]);
        }

        // Handle redirect promise
        await instance.handleRedirectPromise().then((response) => {
          if (response && response.account) {
            instance.setActiveAccount(response.account);
            console.log("✅ Authentication successful");
          }
        });

        // Listen to login events
        instance.addEventCallback((event) => {
          if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
            const payload = event.payload as { account?: AccountInfo };
            if (payload.account) {
              instance.setActiveAccount(payload.account);
            }
          }
        });

        setMsalInstance(instance);
        setIsInitialized(true);
      } catch (error) {
        console.error("❌ MSAL initialization failed:", error);
        setIsInitialized(true); // Still set to true to show error
      }
    };

    initializeMsal();
  }, [config]);

  if (!isInitialized || !msalInstance) {
    return (
      <>
        {loadingComponent || (
          <div style={{
            display: "flex",
            minHeight: "100vh",
            alignItems: "center",
            justifyContent: "center"
          }}>
            <div>Initializing authentication...</div>
          </div>
        )}
      </>
    );
  }

  return (
    <MsalProvider instance={msalInstance}>
      <TeamsAuthInnerProvider scopes={scopes}>
        {children}
      </TeamsAuthInnerProvider>
    </MsalProvider>
  );
}

// ============================================================================
// Hook
// ============================================================================

/**
 * Hook to access Teams authentication context
 *
 * Provides authentication state and methods for signing in/out,
 * getting access tokens, and calling the Microsoft Graph API.
 *
 * @throws Error if used outside of TeamsAuthProvider
 *
 * @example
 * ```tsx
 * function MyComponent() {
 *   const { isAuthenticated, user, login, logout } = useTeamsAuth();
 *
 *   if (!isAuthenticated) {
 *     return <button onClick={login}>Sign In</button>;
 *   }
 *
 *   return (
 *     <div>
 *       <h1>Welcome, {user?.name}!</h1>
 *       <button onClick={logout}>Sign Out</button>
 *     </div>
 *   );
 * }
 * ```
 *
 * @example
 * ```tsx
 * // Get access token for custom API calls
 * const { getAccessToken } = useTeamsAuth();
 *
 * async function fetchData() {
 *   const token = await getAccessToken(['Mail.Read']);
 *   // Use token for API calls
 * }
 * ```
 *
 * @example
 * ```tsx
 * // Call Graph API directly
 * const { callGraphAPI } = useTeamsAuth();
 *
 * async function getUserPresence() {
 *   const presence = await callGraphAPI('/me/presence');
 *   console.log('User is:', presence.availability);
 * }
 * ```
 *
 * @returns {TeamsContextValue} Authentication context with state and methods
 */
export function useTeamsAuth(): TeamsContextValue {
  const context = useContext(TeamsContext);

  if (context === undefined) {
    throw new Error("useTeamsAuth must be used within a TeamsAuthProvider");
  }

  return context;
}
