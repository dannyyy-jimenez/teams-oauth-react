"use client";

import { useState } from "react";
import { useTeamsAuth } from "../components/TeamsAuthProvider";

// ============================================================================
// Types
// ============================================================================

/**
 * Represents a Microsoft Teams team
 */
export interface Team {
  /** Unique identifier for the team */
  id: string;
  /** Display name of the team */
  displayName: string;
  /** Optional description of the team */
  description?: string;
  /** When the team was created */
  createdDateTime?: string;
  /** Web URL to access the team */
  webUrl?: string;
  /** Whether the team is archived */
  isArchived?: boolean;
}

/**
 * Represents a channel within a Microsoft Teams team
 */
export interface Channel {
  /** Unique identifier for the channel */
  id: string;
  /** Display name of the channel */
  displayName: string;
  /** Optional description of the channel */
  description?: string;
  /** Email address associated with the channel */
  email?: string;
  /** Web URL to access the channel */
  webUrl?: string;
  /** Type of membership (standard, private, shared) */
  membershipType?: string;
  /** When the channel was created */
  createdDateTime?: string;
}

/**
 * Represents a message in a Teams channel or chat
 */
export interface Message {
  /** Unique identifier for the message */
  id: string;
  /** When the message was created */
  createdDateTime: string;
  /** When the message was last modified */
  lastModifiedDateTime: string;
  /** Optional subject line */
  subject?: string;
  /** Optional summary of the message */
  summary?: string;
  /** Importance level (normal, high, urgent) */
  importance: string;
  /** Information about who sent the message */
  from: {
    user?: {
      displayName: string;
      id: string;
    };
    application?: {
      displayName: string;
      id: string;
    };
  };
  /** The message body */
  body: {
    contentType: string;
    content: string;
  };
  /** Attachments in the message */
  attachments?: unknown[];
  /** User mentions in the message */
  mentions?: unknown[];
  /** Reactions to the message */
  reactions?: unknown[];
}

/**
 * Represents a member of a Microsoft Teams team
 */
export interface TeamMember {
  /** Unique identifier for the member */
  id: string;
  /** Display name of the member */
  displayName: string;
  /** Email address of the member */
  email?: string;
  /** Roles assigned to the member (owner, member, guest) */
  roles?: string[];
  /** User ID of the member */
  userId?: string;
}

/**
 * Represents a Teams chat conversation
 */
export interface Chat {
  /** Unique identifier for the chat */
  id: string;
  /** Optional topic/title of the chat */
  topic?: string;
  /** When the chat was created */
  createdDateTime: string;
  /** When the chat was last updated */
  lastUpdatedDateTime: string;
  /** Type of chat (oneOnOne, group, meeting) */
  chatType: string;
}

/**
 * Return type for the useTeams hook
 * Provides methods for interacting with Microsoft Teams via Graph API
 */
export interface UseTeamsReturn {
  // Teams
  /** Get all teams the current user is a member of */
  getMyTeams: () => Promise<Team[]>;
  /** Get details about a specific team */
  getTeam: (teamId: string) => Promise<Team>;

  // Channels
  /** Get all channels in a team */
  getTeamChannels: (teamId: string) => Promise<Channel[]>;
  /** Get details about a specific channel */
  getChannel: (teamId: string, channelId: string) => Promise<Channel>;

  // Messages
  /** Get messages from a channel. Optionally specify how many to retrieve (default: 50) */
  getChannelMessages: (teamId: string, channelId: string, top?: number) => Promise<Message[]>;
  /** Send a message to a channel */
  sendChannelMessage: (teamId: string, channelId: string, message: string) => Promise<Message>;

  // Members
  /** Get all members of a team */
  getTeamMembers: (teamId: string) => Promise<TeamMember[]>;

  // Chats
  /** Get all chats for the current user */
  getMyChats: () => Promise<Chat[]>;
  /** Get messages from a chat. Optionally specify how many to retrieve (default: 50) */
  getChatMessages: (chatId: string, top?: number) => Promise<Message[]>;
  /** Send a message to a chat */
  sendChatMessage: (chatId: string, message: string) => Promise<Message>;

  // User Profile
  /** Get the current user's profile information */
  getMyProfile: () => Promise<unknown>;

  // Calendar
  /** Get calendar events for the current user. Optionally filter by date range */
  getMyCalendarEvents: (startDate?: string, endDate?: string) => Promise<unknown[]>;

  // Mail
  /** Get mail messages for the current user. Optionally specify how many to retrieve (default: 10) */
  getMyMail: (top?: number) => Promise<unknown[]>;

  // Generic
  /** Make a custom request to any Microsoft Graph API endpoint */
  customRequest: <T = unknown>(endpoint: string, method?: "GET" | "POST" | "PATCH" | "DELETE", body?: unknown) => Promise<T>;

  // State
  /** Whether an API request is currently in progress */
  loading: boolean;
  /** Current error message, null if no error */
  error: string | null;
  /** Clear the current error message */
  clearError: () => void;
}

// ============================================================================
// Hook
// ============================================================================

/**
 * Hook to access Microsoft Teams API methods
 *
 * Provides a comprehensive set of methods for interacting with Microsoft Teams
 * via the Microsoft Graph API. Handles loading states, errors, and token management.
 *
 * @throws Error if used outside of TeamsAuthProvider
 *
 * @example
 * ```tsx
 * // Basic usage - Get teams and channels
 * function TeamsListComponent() {
 *   const { getMyTeams, getTeamChannels, loading, error } = useTeams();
 *   const [teams, setTeams] = useState([]);
 *
 *   useEffect(() => {
 *     async function loadTeams() {
 *       const myTeams = await getMyTeams();
 *       setTeams(myTeams);
 *     }
 *     loadTeams();
 *   }, []);
 *
 *   if (loading) return <div>Loading...</div>;
 *   if (error) return <div>Error: {error}</div>;
 *
 *   return (
 *     <ul>
 *       {teams.map(team => (
 *         <li key={team.id}>{team.displayName}</li>
 *       ))}
 *     </ul>
 *   );
 * }
 * ```
 *
 * @example
 * ```tsx
 * // Send a message to a Teams channel
 * function SendMessageComponent() {
 *   const { sendChannelMessage, loading, error } = useTeams();
 *
 *   const handleSend = async () => {
 *     await sendChannelMessage(
 *       'team-id',
 *       'channel-id',
 *       'Hello from my app!'
 *     );
 *   };
 *
 *   return (
 *     <button onClick={handleSend} disabled={loading}>
 *       Send Message
 *     </button>
 *   );
 * }
 * ```
 *
 * @example
 * ```tsx
 * // Get and display channel messages
 * function MessagesComponent({ teamId, channelId }) {
 *   const { getChannelMessages } = useTeams();
 *   const [messages, setMessages] = useState([]);
 *
 *   useEffect(() => {
 *     async function loadMessages() {
 *       const msgs = await getChannelMessages(teamId, channelId, 20);
 *       setMessages(msgs);
 *     }
 *     loadMessages();
 *   }, [teamId, channelId]);
 *
 *   return (
 *     <div>
 *       {messages.map(msg => (
 *         <div key={msg.id}>
 *           <strong>{msg.from.user?.displayName}:</strong>
 *           <p>{msg.body.content}</p>
 *         </div>
 *       ))}
 *     </div>
 *   );
 * }
 * ```
 *
 * @example
 * ```tsx
 * // Make a custom API call
 * function UserPresenceComponent() {
 *   const { customRequest } = useTeams();
 *   const [presence, setPresence] = useState(null);
 *
 *   useEffect(() => {
 *     async function getPresence() {
 *       const data = await customRequest('/me/presence', 'GET');
 *       setPresence(data);
 *     }
 *     getPresence();
 *   }, []);
 *
 *   return <div>Status: {presence?.availability}</div>;
 * }
 * ```
 *
 * @returns {UseTeamsReturn} Object containing all Teams API methods, loading state, and error handling
 */
export function useTeams(): UseTeamsReturn {
  const { callGraphAPI, error: authError, clearError: clearAuthError } = useTeamsAuth();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleError = (err: unknown) => {
    const errorMessage = (err as { message?: string }).message || "An error occurred";
    setError(errorMessage);
    console.error("Teams API Error:", err);
    throw err;
  };

  const clearError = () => {
    setError(null);
    clearAuthError();
  };

  // ============================================================================
  // Teams Methods
  // ============================================================================

  const getMyTeams = async (): Promise<Team[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: Team[] }>("me/joinedTeams");
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  const getTeam = async (teamId: string): Promise<Team> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<Team>(`teams/${teamId}`);
      return response;
    } catch (err) {
      handleError(err);
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Channels Methods
  // ============================================================================

  const getTeamChannels = async (teamId: string): Promise<Channel[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: Channel[] }>(`teams/${teamId}/channels`);
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  const getChannel = async (teamId: string, channelId: string): Promise<Channel> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<Channel>(`teams/${teamId}/channels/${channelId}`);
      return response;
    } catch (err) {
      handleError(err);
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Messages Methods
  // ============================================================================

  const getChannelMessages = async (teamId: string, channelId: string, top: number = 50): Promise<Message[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: Message[] }>(
        `teams/${teamId}/channels/${channelId}/messages?$top=${top}`
      );
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  const sendChannelMessage = async (teamId: string, channelId: string, message: string): Promise<Message> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<Message>(
        `teams/${teamId}/channels/${channelId}/messages`,
        "POST",
        {
          body: {
            content: message,
          },
        }
      );
      return response;
    } catch (err) {
      handleError(err);
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Members Methods
  // ============================================================================

  const getTeamMembers = async (teamId: string): Promise<TeamMember[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: TeamMember[] }>(`teams/${teamId}/members`);
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Chats Methods
  // ============================================================================

  const getMyChats = async (): Promise<Chat[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: Chat[] }>("me/chats");
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  const getChatMessages = async (chatId: string, top: number = 50): Promise<Message[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: Message[] }>(
        `chats/${chatId}/messages?$top=${top}`
      );
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  const sendChatMessage = async (chatId: string, message: string): Promise<Message> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<Message>(
        `chats/${chatId}/messages`,
        "POST",
        {
          body: {
            content: message,
          },
        }
      );
      return response;
    } catch (err) {
      handleError(err);
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // User Profile
  // ============================================================================

  const getMyProfile = async (): Promise<unknown> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI("me");
      return response;
    } catch (err) {
      handleError(err);
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Calendar
  // ============================================================================

  const getMyCalendarEvents = async (startDate?: string, endDate?: string): Promise<unknown[]> => {
    setLoading(true);
    setError(null);
    try {
      let endpoint = "me/calendar/events";
      if (startDate && endDate) {
        endpoint += `?startDateTime=${startDate}&endDateTime=${endDate}`;
      }
      const response = await callGraphAPI<{ value: unknown[] }>(endpoint);
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Mail
  // ============================================================================

  const getMyMail = async (top: number = 10): Promise<unknown[]> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<{ value: unknown[] }>(`me/messages?$top=${top}`);
      return response.value || [];
    } catch (err) {
      handleError(err);
      return [];
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Custom Request
  // ============================================================================

  const customRequest = async <T = unknown,>(
    endpoint: string,
    method: "GET" | "POST" | "PATCH" | "DELETE" = "GET",
    body?: unknown
  ): Promise<T> => {
    setLoading(true);
    setError(null);
    try {
      const response = await callGraphAPI<T>(endpoint, method, body);
      return response;
    } catch (err) {
      handleError(err);
      throw err;
    } finally {
      setLoading(false);
    }
  };

  // ============================================================================
  // Return
  // ============================================================================

  return {
    getMyTeams,
    getTeam,
    getTeamChannels,
    getChannel,
    getChannelMessages,
    sendChannelMessage,
    getTeamMembers,
    getMyChats,
    getChatMessages,
    sendChatMessage,
    getMyProfile,
    getMyCalendarEvents,
    getMyMail,
    customRequest,
    loading,
    error: error || authError,
    clearError,
  };
}
