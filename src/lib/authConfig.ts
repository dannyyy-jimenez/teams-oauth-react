import { Configuration, PopupRequest } from "@azure/msal-browser";

/**
 * Configuration object to be passed to MSAL instance on creation.
 *
 * This configuration is built dynamically by the TeamsAuthProvider based on
 * the config passed to it. You typically don't need to use this directly.
 *
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md
 *
 * @internal
 */
export const msalConfig: Configuration = {
  auth: {
    clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || "", // This is the ONLY mandatory field that you need to supply.
    authority: `https://login.microsoftonline.com/${process.env.NEXT_PUBLIC_AZURE_TENANT_ID || "common"}`, // Use "common" for multi-tenant apps
    redirectUri: process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000", // You must register this URI on Azure Portal/App Registration.
    postLogoutRedirectUri: process.env.NEXT_PUBLIC_REDIRECT_URI || "http://localhost:3000", // Redirect here after logout
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
  },
};

/**
 * Default login request configuration
 *
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 *
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 *
 * @internal
 */
export const loginRequest: PopupRequest = {
  scopes: ["User.Read"], // Basic profile information
};

/**
 * Microsoft Graph API endpoint configuration
 *
 * Common Graph API endpoints used for user and Teams data.
 *
 * For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
  graphUsersEndpoint: "https://graph.microsoft.com/v1.0/users",
};

/**
 * Predefined scope groups for Microsoft Graph API
 *
 * These scope collections make it easy to request the right permissions
 * for different features. Mix and match based on your needs.
 *
 * Teams API scopes reference: https://learn.microsoft.com/en-us/graph/permissions-reference
 *
 * @example
 * ```tsx
 * // Use in TeamsAuthProvider config
 * <TeamsAuthProvider
 *   config={{
 *     clientId: "...",
 *     scopes: graphScopes.teams // or graphScopes.teamsExtended
 *   }}
 * />
 * ```
 */
export const graphScopes = {
  /** Basic user profile scope - read user information */
  profile: ["User.Read"],

  /** Standard Teams scopes - read teams, channels, and messages */
  teams: [
    "User.Read",
    "Team.ReadBasic.All", // Read basic team info
    "Channel.ReadBasic.All", // Read basic channel info
    "ChannelMessage.Read.All", // Read channel messages
  ],

  /** Extended Teams scopes - includes message sending and member management */
  teamsExtended: [
    "User.Read",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Read.All",
    "ChannelMessage.Send", // Send messages to channels
    "TeamMember.Read.All", // Read team members
    "Chat.Read", // Read user's chats
    "Chat.ReadWrite", // Read and write user's chats
  ],

  /** Calendar scopes - read and write calendar events */
  calendar: [
    "User.Read",
    "Calendars.Read",
    "Calendars.ReadWrite",
  ],

  /** Mail scopes - read and send email messages */
  mail: [
    "User.Read",
    "Mail.Read",
    "Mail.Send",
  ],
};

/**
 * Default scopes for initial login
 *
 * These scopes are used by default if no custom scopes are provided
 * to the TeamsAuthProvider. Includes basic Teams read permissions.
 *
 * @example
 * ```tsx
 * // These scopes are automatically used
 * <TeamsAuthProvider config={{ clientId: "..." }} />
 *
 * // Or explicitly override
 * <TeamsAuthProvider
 *   config={{
 *     clientId: "...",
 *     scopes: graphScopes.teamsExtended
 *   }}
 * />
 * ```
 */
export const defaultScopes = graphScopes.teams;
