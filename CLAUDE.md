# CLAUDE.md - AI Assistant Guide

This document helps Claude (or other AI assistants) understand the architecture and design decisions of this Teams OAuth Plugin.

## 🏗️ Architecture Overview

### Purpose
A reusable, drop-in Microsoft Teams OAuth authentication plugin for Next.js applications with full Graph API integration.

### Design Philosophy
1. **Headless** - No UI components, only logic and state management
2. **Simple API** - Two hooks (`useTeamsAuth`, `useTeams`) cover all use cases
3. **Type-Safe** - Full TypeScript support with exported types
4. **Framework-Aware** - Built specifically for Next.js App Router with React Server Components
5. **Production-Ready** - Error handling, loading states, token refresh built-in

## 📂 Project Structure

```
package-teams-oauth/
├── src/
│   ├── components/
│   │   └── TeamsAuthProvider.tsx    # Main provider component with context
│   ├── hooks/
│   │   └── useTeams.ts              # Teams-specific API methods hook
│   ├── lib/
│   │   ├── authConfig.ts            # MSAL configuration and scopes
│   │   └── graphClient.ts           # Graph API helper functions
│   └── index.ts                      # Public API exports
├── package.json                       # Package configuration
├── tsconfig.json                      # TypeScript configuration
├── tsup.config.ts                     # Build configuration (ESM + CJS)
├── README.md                          # User documentation
├── USAGE_IN_PROJECT.md               # Integration guide
└── CLAUDE.md                          # This file
```

## 🧩 Component Breakdown

### 1. TeamsAuthProvider.tsx
**Purpose**: Manages MSAL initialization and provides authentication context

**Key Responsibilities**:
- Initialize MSAL PublicClientApplication
- Handle OAuth redirect flow (not popup - more reliable)
- Provide authentication state via React Context
- Expose auth methods (login, logout, getAccessToken)
- Provide callGraphAPI helper for custom API calls

**Architecture Pattern**: Provider + Context + Hook
- Outer provider: Initializes MSAL, wraps with MsalProvider
- Inner provider: Uses MSAL hooks, creates Teams context
- Hook: Consumes context, ensures it's used correctly

**Why Two Providers?**:
- Outer: Can't use MSAL hooks (not inside MsalProvider yet)
- Inner: Can use MSAL hooks (inside MsalProvider now)
- Separation allows clean initialization and hook usage

**Important Details**:
- Uses redirect flow, NOT popup (avoids popup blockers)
- Handles `handleRedirectPromise()` on mount
- Sets active account automatically
- Listens to LOGIN_SUCCESS event

### 2. useTeams.ts
**Purpose**: Provides convenient methods for common Teams operations

**Design Decision**: Separate from auth hook
- `useTeamsAuth()` - Authentication primitives
- `useTeams()` - Business logic for Teams API

**Why Separate?**:
- Separation of concerns
- Users might want auth without Teams methods
- Easier to add more domain-specific hooks (useCalendar, useMail, etc.)
- Cleaner API surface

**Pattern**: Thin wrapper around callGraphAPI
- Each method is a specialized API call
- Handles loading and error states internally
- Returns typed data structures
- Consistent error handling

### 3. authConfig.ts
**Purpose**: Centralize MSAL configuration and permission scopes

**Key Exports**:
- `msalConfig` - MSAL configuration template (internal)
- `graphScopes` - Predefined scope collections
- `defaultScopes` - Default scopes if none provided

**Design Decision**: Scope Collections
Instead of requiring users to remember scope strings, we provide:
```typescript
graphScopes.teams          // Basic read
graphScopes.teamsExtended  // Read + write
graphScopes.calendar       // Calendar access
graphScopes.mail           // Mail access
```

### 4. graphClient.ts
**Purpose**: Low-level Graph API helpers (optional, for advanced users)

**Why Include?**:
- Some users prefer direct Graph client access
- Useful for features not covered by useTeams hook
- Provides graphAPI namespace with helper functions
- Legacy/backward compatibility

**Pattern**: Factory + Namespace
- `getGraphClient()` - Creates authenticated client
- `graphAPI.*` - Collection of helper methods
- `handleGraphError()` - User-friendly error messages

## 🔑 Key Design Decisions

### 1. Redirect Flow Over Popup
**Decision**: Use `loginRedirect` instead of `loginPopup`

**Reasoning**:
- Popup blockers are common
- Mobile devices handle redirects better
- More reliable cross-browser
- Industry standard for SPAs

**Trade-off**:
- Page refresh after login (acceptable)
- Can't have modal-style login (rarely needed)

### 2. React Context for State
**Decision**: Use React Context instead of external state management

**Reasoning**:
- Simple, built-in solution
- No external dependencies
- Sufficient for auth state (not complex)
- Familiar to React developers

**Trade-off**:
- Re-renders on any context change (mitigated by splitting context)
- Not suitable for very large apps (can be refactored if needed)

### 3. TypeScript First
**Decision**: Written in TypeScript with full type exports

**Reasoning**:
- Better DX with autocomplete
- Catches errors at compile time
- Self-documenting API
- Industry expectation for libraries

### 4. Headless Design
**Decision**: No UI components included

**Reasoning**:
- Maximum flexibility for users
- Doesn't impose design decisions
- Smaller bundle size
- Easier to maintain

**Trade-off**:
- Users must build UI (but we provide examples)
- More setup for simple cases (mitigated by good docs)

### 5. Next.js App Router Focus
**Decision**: Built for Next.js 13+ App Router

**Reasoning**:
- Modern Next.js apps use App Router
- `"use client"` directive for client components
- Server Components compatible (provider in client component)

**Compatibility**:
- Still works with Pages Router
- Works with plain React (just needs React 18+)

## 🔄 Data Flow

### Authentication Flow
```
1. User clicks login button
   ↓
2. useTeamsAuth().login() called
   ↓
3. MSAL redirects to Microsoft login
   ↓
4. User authenticates
   ↓
5. Microsoft redirects back to app
   ↓
6. handleRedirectPromise() processes response
   ↓
7. Context updates with user info
   ↓
8. Components re-render with isAuthenticated=true
```

### API Call Flow
```
1. Component calls useTeams().getMyTeams()
   ↓
2. Hook calls callGraphAPI() from useTeamsAuth()
   ↓
3. callGraphAPI gets access token via getAccessToken()
   ↓
4. If token expired, acquireTokenSilent() refreshes it
   ↓
5. Make fetch request with Bearer token
   ↓
6. Parse response
   ↓
7. Return typed data to component
```

## 🛠️ Common Modifications

### Adding a New Teams API Method

1. **Add to useTeams.ts**:
```typescript
const getTeamSettings = async (teamId: string): Promise<TeamSettings> => {
  setLoading(true);
  setError(null);
  try {
    const response = await callGraphAPI<TeamSettings>(`teams/${teamId}/settings`);
    return response;
  } catch (err) {
    handleError(err);
    throw err;
  } finally {
    setLoading(false);
  }
};
```

2. **Add to UseTeamsReturn interface**:
```typescript
export interface UseTeamsReturn {
  // ... existing methods
  getTeamSettings: (teamId: string) => Promise<TeamSettings>;
}
```

3. **Return in hook**:
```typescript
return {
  // ... existing methods
  getTeamSettings,
};
```

### Adding a New Permission Scope Group

Add to `authConfig.ts`:
```typescript
export const graphScopes = {
  // ... existing scopes
  myCustomScopes: [
    "User.Read",
    "Files.Read.All",
    "Sites.Read.All",
  ],
};
```

### Adding Error Handling for Specific Error

Add to `graphClient.ts`:
```typescript
export function handleGraphError(error: any): string {
  if (error.statusCode === 429) {
    return "Too many requests. Please try again later.";
  }
  // ... existing error handlers
}
```

## 🧪 Testing Considerations

### What to Test
1. **Provider initialization**
   - MSAL initializes correctly
   - Handles redirect promise
   - Sets active account

2. **Auth methods**
   - Login redirects correctly
   - Logout clears state
   - Token acquisition works

3. **Teams methods**
   - Each method calls correct endpoint
   - Loading states work
   - Errors are handled

4. **Error scenarios**
   - Token expired
   - Network failure
   - Permission denied
   - Invalid IDs

### Mocking Strategy
```typescript
// Mock MSAL
jest.mock('@azure/msal-react', () => ({
  useMsal: () => ({
    instance: mockInstance,
    accounts: mockAccounts,
    inProgress: 'none',
  }),
  useIsAuthenticated: () => true,
}));

// Mock Graph API
global.fetch = jest.fn(() =>
  Promise.resolve({
    ok: true,
    json: () => Promise.resolve({ value: [] }),
  })
);
```

## 📝 Documentation Standards

### JSDoc Format
```typescript
/**
 * Brief one-line description
 *
 * Longer description with context and details.
 * Can be multiple paragraphs.
 *
 * @param paramName - Description of parameter
 * @returns Description of return value
 * @throws Description of errors thrown
 *
 * @example
 * ```typescript
 * const result = await someFunction('example');
 * console.log(result);
 * ```
 *
 * @see RelatedFunction
 */
```

### README Structure
1. Quick start example (copy-paste ready)
2. API reference (all exports documented)
3. Examples (common use cases)
4. Setup guide (Azure AD configuration)
5. Troubleshooting (common issues)

## 🚀 Build & Distribution

### Build Process (tsup)
- Input: `src/**/*.ts(x)`
- Output: `dist/` with:
  - ESM (`.mjs`)
  - CommonJS (`.js`)
  - Type definitions (`.d.ts`)
  - Source maps

### Package Exports
```json
{
  ".": "Main exports (provider + hooks)",
  "./components": "Just the provider",
  "./hooks": "Just the hooks"
}
```

### Peer Dependencies
- React 18+
- Next.js 13+
- `@azure/msal-browser`
- `@azure/msal-react`
- `@microsoft/microsoft-graph-client`

## 🔐 Security Considerations

### Token Storage
- Stored in sessionStorage by default
- Can be changed to localStorage (less secure but persists)
- Never exposed in URL or logs

### Scope Principle
- Only request necessary permissions
- Users can customize scopes
- Admin consent may be required

### Redirect URI Validation
- Must be registered in Azure AD
- Exact match required (including protocol, domain, port, path)
- localhost allowed for development

## 🐛 Common Issues & Solutions

### Issue: "No account found"
**Cause**: User not authenticated
**Solution**: Check `isAuthenticated` before calling API methods

### Issue: "Insufficient permissions"
**Cause**: Missing scope in Azure AD
**Solution**: Add required permission and grant consent

### Issue: "redirect_uri_mismatch"
**Cause**: Redirect URI doesn't match Azure AD config
**Solution**: Ensure exact match (check port, protocol, trailing slash)

### Issue: Token refresh fails
**Cause**: Refresh token expired (usually after 90 days)
**Solution**: User must log in again (automatic via redirect)

## 💡 Future Enhancements

Potential additions without breaking changes:

1. **Additional Hooks**
   - `useCalendar()` - Calendar-specific operations
   - `useMail()` - Mail-specific operations
   - `useFiles()` - SharePoint/OneDrive files

2. **Optimizations**
   - Request batching for multiple API calls
   - Response caching with TTL
   - Retry logic with exponential backoff

3. **Features**
   - Webhook subscriptions for real-time updates
   - File upload/download helpers
   - Presence status helpers

4. **Developer Experience**
   - CLI tool for setup
   - Vite plugin
   - Testing utilities

## 📚 Resources

- [Microsoft Graph API Docs](https://learn.microsoft.com/en-us/graph/)
- [MSAL.js Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [Azure AD App Registration](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Teams API Reference](https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview)

## 🤝 Contributing Guidelines

When modifying this package:

1. **Maintain backward compatibility** - Don't break existing APIs
2. **Add JSDoc comments** - All public APIs must be documented
3. **Update README** - Document new features
4. **Add examples** - Show how to use new features
5. **Test thoroughly** - Ensure nothing breaks
6. **Update CLAUDE.md** - Explain design decisions

## 📄 License

MIT - See LICENSE file for details

---

**For AI Assistants**: This file provides context about the codebase architecture, design decisions, and common patterns. Use it to understand the system before making modifications or answering questions about the code.
