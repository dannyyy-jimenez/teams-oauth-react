# 🚀 Teams OAuth Plugin

A ready-to-use Microsoft Teams OAuth plugin for Next.js with full Graph API integration.

## ✨ Features

- 🔐 **Easy OAuth Authentication** - Drop-in provider and hook
- 🎯 **Teams API Methods** - Pre-built methods for common operations
- 📦 **TypeScript Support** - Fully typed
- ⚡ **Next.js 13+ Compatible** - Works with App Router
- 🔄 **Auto Token Refresh** - Handles token management automatically
- 🎨 **Headless** - Bring your own UI

## 📦 Installation

```bash
npm install @your-org/teams-oauth-plugin
# or
yarn add @your-org/teams-oauth-plugin
# or
pnpm add @your-org/teams-oauth-plugin
```

## 🚀 Quick Start

### 1. Wrap your app with the provider

```tsx
// app/layout.tsx
import { TeamsAuthProvider } from '@your-org/teams-oauth-plugin';

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html>
      <body>
        <TeamsAuthProvider
          config={{
            clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID!,
            tenantId: process.env.NEXT_PUBLIC_AZURE_TENANT_ID, // optional, defaults to "common"
            redirectUri: process.env.NEXT_PUBLIC_REDIRECT_URI, // optional, defaults to window.location.origin
            scopes: [ // optional, these are the defaults
              'User.Read',
              'Team.ReadBasic.All',
              'Channel.ReadBasic.All',
              'ChannelMessage.Read.All',
            ],
          }}
        >
          {children}
        </TeamsAuthProvider>
      </body>
    </html>
  );
}
```

### 2. Use the hooks in your components

```tsx
'use client';

import { useTeamsAuth, useTeams } from '@your-org/teams-oauth-plugin';

export default function MyComponent() {
  // Auth hook - for authentication state and methods
  const { isAuthenticated, user, login, logout, isLoading } = useTeamsAuth();

  // Teams hook - for Teams API methods
  const { getMyTeams, getTeamChannels, sendChannelMessage, loading, error } = useTeams();

  const handleGetTeams = async () => {
    const teams = await getMyTeams();
    console.log('My teams:', teams);
  };

  if (!isAuthenticated) {
    return <button onClick={login}>Sign in with Microsoft</button>;
  }

  return (
    <div>
      <h1>Welcome, {user?.name}!</h1>
      <button onClick={handleGetTeams} disabled={loading}>
        Get My Teams
      </button>
      <button onClick={logout}>Sign Out</button>
      {error && <p>Error: {error}</p>}
    </div>
  );
}
```

## 📚 API Reference

### `<TeamsAuthProvider>`

Wraps your app and provides authentication context.

**Props:**
- `config: TeamsAuthConfig` - Configuration object
  - `clientId: string` - Azure AD App Client ID (required)
  - `tenantId?: string` - Azure AD Tenant ID (optional, defaults to "common")
  - `redirectUri?: string` - OAuth redirect URI (optional, defaults to current origin)
  - `scopes?: string[]` - Permission scopes (optional, has sensible defaults)
  - `cacheLocation?: 'localStorage' | 'sessionStorage'` - Token cache location (optional, defaults to sessionStorage)
- `children: ReactNode` - Your app content
- `loadingComponent?: ReactNode` - Custom loading component (optional)

### `useTeamsAuth()`

Hook for authentication state and methods.

**Returns:**
```typescript
{
  isAuthenticated: boolean;           // Whether user is signed in
  isLoading: boolean;                 // Whether auth is in progress
  user: AccountInfo | null;           // Current user info
  inProgress: InteractionStatus;      // Current MSAL interaction status
  login: () => Promise<void>;         // Sign in method
  logout: () => Promise<void>;        // Sign out method
  getAccessToken: (scopes?: string[]) => Promise<string | null>; // Get access token
  callGraphAPI: <T>(endpoint, method?, body?) => Promise<T>; // Call any Graph API endpoint
  error: string | null;               // Error message
  clearError: () => void;             // Clear error
}
```

### `useTeams()`

Hook for Teams-specific API methods.

**Returns:**
```typescript
{
  // Teams
  getMyTeams: () => Promise<Team[]>;
  getTeam: (teamId: string) => Promise<Team>;

  // Channels
  getTeamChannels: (teamId: string) => Promise<Channel[]>;
  getChannel: (teamId: string, channelId: string) => Promise<Channel>;

  // Messages
  getChannelMessages: (teamId: string, channelId: string, top?: number) => Promise<Message[]>;
  sendChannelMessage: (teamId: string, channelId: string, message: string) => Promise<Message>;

  // Members
  getTeamMembers: (teamId: string) => Promise<TeamMember[]>;

  // Chats
  getMyChats: () => Promise<Chat[]>;
  getChatMessages: (chatId: string, top?: number) => Promise<Message[]>;
  sendChatMessage: (chatId: string, message: string) => Promise<Message>;

  // User
  getMyProfile: () => Promise<any>;

  // Calendar
  getMyCalendarEvents: (startDate?: string, endDate?: string) => Promise<any[]>;

  // Mail
  getMyMail: (top?: number) => Promise<any[]>;

  // Generic
  customRequest: <T>(endpoint: string, method?, body?) => Promise<T>;

  // State
  loading: boolean;
  error: string | null;
  clearError: () => void;
}
```

## 🎯 Examples

### Get Teams and Channels

```tsx
const { getMyTeams, getTeamChannels } = useTeams();

const [teams, setTeams] = useState<Team[]>([]);
const [channels, setChannels] = useState<Channel[]>([]);

useEffect(() => {
  async function loadTeams() {
    const myTeams = await getMyTeams();
    setTeams(myTeams);
  }
  loadTeams();
}, []);

const handleTeamClick = async (teamId: string) => {
  const teamChannels = await getTeamChannels(teamId);
  setChannels(teamChannels);
};
```

### Send a Message

```tsx
const { sendChannelMessage, error } = useTeams();

const handleSendMessage = async () => {
  try {
    await sendChannelMessage(
      'team-id',
      'channel-id',
      'Hello from my app!'
    );
    alert('Message sent!');
  } catch (err) {
    console.error('Failed to send:', error);
  }
};
```

### Custom API Call

```tsx
const { customRequest } = useTeams();

const getUserPresence = async () => {
  const presence = await customRequest('/me/presence', 'GET');
  console.log('User is:', presence.availability);
};
```

### Conditional Rendering

```tsx
const { isAuthenticated, isLoading, login } = useTeamsAuth();

if (isLoading) {
  return <div>Loading...</div>;
}

if (!isAuthenticated) {
  return <button onClick={login}>Sign In</button>;
}

return <YourApp />;
```

## 🔧 Azure AD Setup

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Create a new registration:
   - Name: Your app name
   - Supported account types: Choose based on your needs
   - Redirect URI: **Single-page application** → `http://localhost:3000`
4. Note your **Application (client) ID** and **Directory (tenant) ID**
5. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated**
6. Add these permissions:
   - `User.Read`
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `ChannelMessage.Read.All`
   - `ChannelMessage.Send` (if sending messages)
7. **Grant admin consent** (if you're an admin)

## 📝 Environment Variables

Create a `.env.local` file:

```env
NEXT_PUBLIC_AZURE_CLIENT_ID=your-client-id
NEXT_PUBLIC_AZURE_TENANT_ID=your-tenant-id-or-common
NEXT_PUBLIC_REDIRECT_URI=http://localhost:3000
```

## 🔒 Security

- Never commit `.env.local` to version control
- Use HTTPS in production
- Only request the permissions you need
- Tokens are stored in sessionStorage by default (can be changed to localStorage)

## 📖 TypeScript Support

All hooks and components are fully typed. Import types:

```typescript
import type {
  TeamsAuthConfig,
  Team,
  Channel,
  Message,
  AccountInfo,
} from '@your-org/teams-oauth-plugin';
```

## 🐛 Troubleshooting

### "No account found" error
- Make sure you're signed in first
- Check if `isAuthenticated` is true

### "Insufficient permissions" error
- Add the required permission in Azure AD
- Grant admin consent if needed

### Popup blocked
- The plugin uses redirect flow by default (not popup)
- No popup blockers should interfere

### Token expired
- Tokens are automatically refreshed
- If refresh fails, user will be redirected to sign in again

## 📄 License

MIT

## 🤝 Contributing

Contributions welcome! Please open an issue or PR.

## 🔗 Links

- [Microsoft Graph API Docs](https://learn.microsoft.com/en-us/graph/overview)
- [Teams API Reference](https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview)
- [Azure AD App Registration](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
