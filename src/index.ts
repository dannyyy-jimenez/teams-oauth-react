/**
 * @packageDocumentation
 *
 * Microsoft Teams OAuth Plugin for Next.js
 *
 * A production-ready authentication solution for Microsoft Teams with full Graph API integration.
 *
 * @example
 * ```tsx
 * // 1. Wrap your app with the provider
 * import { TeamsAuthProvider } from '@your-org/teams-oauth-plugin';
 *
 * export default function RootLayout({ children }) {
 *   return (
 *     <TeamsAuthProvider config={{ clientId: "your-client-id" }}>
 *       {children}
 *     </TeamsAuthProvider>
 *   );
 * }
 *
 * // 2. Use the hooks in your components
 * import { useTeamsAuth, useTeams } from '@your-org/teams-oauth-plugin';
 *
 * function MyComponent() {
 *   const { isAuthenticated, login, logout } = useTeamsAuth();
 *   const { getMyTeams, loading } = useTeams();
 *
 *   if (!isAuthenticated) {
 *     return <button onClick={login}>Sign In</button>;
 *   }
 *
 *   return <button onClick={logout}>Sign Out</button>;
 * }
 * ```
 */

// ============================================================================
// Main Provider and Authentication Hook
// ============================================================================

/**
 * Main authentication provider component and hook for Microsoft Teams OAuth
 *
 * @module TeamsAuth
 */
export { TeamsAuthProvider, useTeamsAuth } from "./components/TeamsAuthProvider";

/**
 * TypeScript types for authentication configuration and context
 *
 * @module TeamsAuth.Types
 */
export type { TeamsAuthConfig, TeamsContextValue } from "./components/TeamsAuthProvider";

// ============================================================================
// Teams API Hook
// ============================================================================

/**
 * Hook for accessing Microsoft Teams API methods
 * Provides convenient methods for teams, channels, messages, chats, and more
 *
 * @module TeamsAPI
 */
export { useTeams } from "./hooks/useTeams";

/**
 * TypeScript types for Teams API data structures
 *
 * @module TeamsAPI.Types
 */
export type {
  UseTeamsReturn,
  Team,
  Channel,
  Message,
  TeamMember,
  Chat,
} from "./hooks/useTeams";

// ============================================================================
// Library Exports (Advanced Usage)
// ============================================================================

/**
 * MSAL configuration and permission scopes
 * For advanced users who need direct access to auth configuration
 *
 * @module AuthConfig
 */
export * from "./lib/authConfig";

/**
 * Microsoft Graph API client helpers
 * For advanced users who need direct Graph API access
 *
 * @module GraphClient
 */
export * from "./lib/graphClient";

// ============================================================================
// MSAL Type Re-exports
// ============================================================================

/**
 * Commonly used MSAL types re-exported for convenience
 *
 * @module MSAL.Types
 */
export type { AccountInfo, InteractionStatus } from "@azure/msal-browser";
