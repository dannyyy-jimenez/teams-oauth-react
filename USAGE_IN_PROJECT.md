# 📦 How to Use This Plugin in Your Project

## Option 1: Local Package (Development/Testing)

### Step 1: Build the package

```bash
cd package-teams-oauth
npm install
npm run build
```

### Step 2: Link it locally

```bash
# In the package directory
npm link

# In your Next.js project
npm link @your-org/teams-oauth-plugin
```

### Step 3: Use in your project

```tsx
// app/layout.tsx
import { TeamsAuthProvider } from '@your-org/teams-oauth-plugin';

export default function RootLayout({ children }) {
  return (
    <html>
      <body>
        <TeamsAuthProvider
          config={{
            clientId: process.env.NEXT_PUBLIC_AZURE_CLIENT_ID!,
            tenantId: process.env.NEXT_PUBLIC_AZURE_TENANT_ID,
          }}
        >
          {children}
        </TeamsAuthProvider>
      </body>
    </html>
  );
}
```

```tsx
// app/page.tsx
'use client';

import { useTeamsAuth, useTeams } from '@your-org/teams-oauth-plugin';

export default function Home() {
  const { isAuthenticated, user, login, logout } = useTeamsAuth();
  const { getMyTeams, loading } = useTeams();

  const handleGetTeams = async () => {
    const teams = await getMyTeams();
    console.log(teams);
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
    </div>
  );
}
```

## Option 2: Copy Into Project Directly

### Step 1: Copy the source files

```bash
# From package-teams-oauth/src directory, copy to your project:
cp -r src/components your-project/lib/teams-oauth/
cp -r src/hooks your-project/lib/teams-oauth/
cp -r src/lib your-project/lib/teams-oauth/
```

### Step 2: Update imports

Change from:
```tsx
import { TeamsAuthProvider, useTeams } from '@your-org/teams-oauth-plugin';
```

To:
```tsx
import { TeamsAuthProvider } from '@/lib/teams-oauth/components/TeamsAuthProvider';
import { useTeams } from '@/lib/teams-oauth/hooks/useTeams';
```

## Option 3: Publish to NPM (Production)

### Step 1: Update package.json

```bash
cd package-teams-oauth
# Update name, author, repository in package.json
```

### Step 2: Publish

```bash
npm login
npm publish --access public
```

### Step 3: Install in any project

```bash
npm install @your-org/teams-oauth-plugin
```

## 🎯 Complete Example

Here's a full example showing all the features:

```tsx
// app/teams-dashboard/page.tsx
'use client';

import { useState, useEffect } from 'react';
import { useTeamsAuth, useTeams, type Team, type Channel } from '@your-org/teams-oauth-plugin';

export default function TeamsDashboard() {
  const { isAuthenticated, user, login, logout, isLoading } = useTeamsAuth();
  const {
    getMyTeams,
    getTeamChannels,
    getChannelMessages,
    sendChannelMessage,
    loading,
    error,
    clearError,
  } = useTeams();

  const [teams, setTeams] = useState<Team[]>([]);
  const [selectedTeam, setSelectedTeam] = useState<string | null>(null);
  const [channels, setChannels] = useState<Channel[]>([]);

  // Load teams on mount
  useEffect(() => {
    if (isAuthenticated) {
      loadTeams();
    }
  }, [isAuthenticated]);

  const loadTeams = async () => {
    try {
      const myTeams = await getMyTeams();
      setTeams(myTeams);
    } catch (err) {
      console.error('Failed to load teams:', err);
    }
  };

  const handleTeamClick = async (teamId: string) => {
    setSelectedTeam(teamId);
    try {
      const teamChannels = await getTeamChannels(teamId);
      setChannels(teamChannels);
    } catch (err) {
      console.error('Failed to load channels:', err);
    }
  };

  const handleSendMessage = async (channelId: string) => {
    if (!selectedTeam) return;

    try {
      await sendChannelMessage(
        selectedTeam,
        channelId,
        'Hello from my custom app! 🚀'
      );
      alert('Message sent successfully!');
    } catch (err) {
      console.error('Failed to send message:', err);
    }
  };

  // Loading state
  if (isLoading) {
    return <div>Loading...</div>;
  }

  // Not authenticated
  if (!isAuthenticated) {
    return (
      <div style={{ textAlign: 'center', padding: '50px' }}>
        <h1>Microsoft Teams Dashboard</h1>
        <button onClick={login} style={{ padding: '10px 20px', fontSize: '16px' }}>
          Sign in with Microsoft
        </button>
      </div>
    );
  }

  // Authenticated
  return (
    <div style={{ padding: '20px' }}>
      <header style={{ marginBottom: '20px' }}>
        <h1>Welcome, {user?.name}!</h1>
        <p>Email: {user?.username}</p>
        <button onClick={logout}>Sign Out</button>
      </header>

      {error && (
        <div style={{ background: '#fee', padding: '10px', marginBottom: '20px' }}>
          <strong>Error:</strong> {error}
          <button onClick={clearError} style={{ marginLeft: '10px' }}>✕</button>
        </div>
      )}

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '20px' }}>
        {/* Teams List */}
        <div>
          <h2>Your Teams ({teams.length})</h2>
          {loading && <p>Loading teams...</p>}
          {teams.map((team) => (
            <div
              key={team.id}
              onClick={() => handleTeamClick(team.id)}
              style={{
                padding: '10px',
                border: '1px solid #ccc',
                marginBottom: '10px',
                cursor: 'pointer',
                background: selectedTeam === team.id ? '#e3f2fd' : 'white',
              }}
            >
              <strong>{team.displayName}</strong>
              {team.description && <p>{team.description}</p>}
            </div>
          ))}
        </div>

        {/* Channels List */}
        <div>
          <h2>Channels {selectedTeam ? `(${channels.length})` : ''}</h2>
          {selectedTeam ? (
            <>
              {loading && <p>Loading channels...</p>}
              {channels.map((channel) => (
                <div
                  key={channel.id}
                  style={{
                    padding: '10px',
                    border: '1px solid #ccc',
                    marginBottom: '10px',
                  }}
                >
                  <strong># {channel.displayName}</strong>
                  {channel.description && <p>{channel.description}</p>}
                  <button
                    onClick={() => handleSendMessage(channel.id)}
                    style={{ marginTop: '5px' }}
                  >
                    Send Message
                  </button>
                </div>
              ))}
            </>
          ) : (
            <p>Select a team to view channels</p>
          )}
        </div>
      </div>
    </div>
  );
}
```

## 🔑 Environment Variables

Don't forget to create `.env.local`:

```env
NEXT_PUBLIC_AZURE_CLIENT_ID=your-client-id-here
NEXT_PUBLIC_AZURE_TENANT_ID=your-tenant-id-or-common
NEXT_PUBLIC_REDIRECT_URI=http://localhost:3000
```

## 🚀 Next Steps

1. ✅ Choose installation method (link, copy, or publish)
2. ✅ Set up Azure AD app registration
3. ✅ Add environment variables
4. ✅ Wrap app with `<TeamsAuthProvider>`
5. ✅ Use `useTeamsAuth()` and `useTeams()` hooks
6. ✅ Build your app!

## 💡 Pro Tips

- Use `useTeamsAuth()` for auth state/methods
- Use `useTeams()` for Teams API calls
- Check `loading` state before showing results
- Handle `error` state gracefully
- Use TypeScript types for better DX
- Test with real Teams data
