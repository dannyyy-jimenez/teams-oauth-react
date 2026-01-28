import { Client } from "@microsoft/microsoft-graph-client";

/**
 * Creates an authenticated Microsoft Graph client
 *
 * This function initializes a Graph API client with the provided access token.
 * The client handles authentication automatically for all subsequent API calls.
 *
 * @param accessToken - The access token obtained from MSAL
 * @returns Configured Microsoft Graph client instance
 *
 * @example
 * ```typescript
 * const token = await getAccessToken();
 * const client = getGraphClient(token);
 * const user = await client.api('/me').get();
 * ```
 */
export function getGraphClient(accessToken: string): Client {
  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    },
  });
}

/**
 * Collection of helper functions for Microsoft Graph API
 *
 * Provides convenient methods for common Teams operations including
 * teams, channels, messages, chats, calendar, and mail.
 *
 * @namespace graphAPI
 */
export const graphAPI = {
  /**
   * Get the current user's profile information
   *
   * @param accessToken - Valid access token with User.Read scope
   * @returns User profile object containing displayName, email, jobTitle, etc.
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const profile = await graphAPI.getMyProfile(token);
   * console.log(profile.displayName, profile.mail);
   * ```
   */
  async getMyProfile(accessToken: string) {
    const client = getGraphClient(accessToken);
    return await client.api("/me").get();
  },

  /**
   * Get the current user's profile photo
   *
   * @param accessToken - Valid access token with User.Read scope
   * @returns Photo blob or null if no photo is found
   *
   * @example
   * ```typescript
   * const photo = await graphAPI.getMyPhoto(token);
   * if (photo) {
   *   // Display photo
   * }
   * ```
   */
  async getMyPhoto(accessToken: string) {
    const client = getGraphClient(accessToken);
    try {
      const photo = await client.api("/me/photo/$value").get();
      return photo;
    } catch {
      console.log("No profile photo found");
      return null;
    }
  },

  /**
   * Get all Teams the current user is a member of
   *
   * @param accessToken - Valid access token with Team.ReadBasic.All scope
   * @returns Object containing an array of teams in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getMyTeams(token);
   * const teams = response.value;
   * teams.forEach(team => console.log(team.displayName));
   * ```
   */
  async getMyTeams(accessToken: string) {
    const client = getGraphClient(accessToken);
    return await client.api("/me/joinedTeams").get();
  },

  /**
   * Get detailed information about a specific team
   *
   * @param accessToken - Valid access token with Team.ReadBasic.All scope
   * @param teamId - The unique identifier of the team
   * @returns Team object with full details
   * @throws Error if the API call fails or team is not found
   *
   * @example
   * ```typescript
   * const team = await graphAPI.getTeam(token, 'team-id-here');
   * console.log(team.displayName, team.description);
   * ```
   */
  async getTeam(accessToken: string, teamId: string) {
    const client = getGraphClient(accessToken);
    return await client.api(`/teams/${teamId}`).get();
  },

  /**
   * Get all channels in a specific team
   *
   * @param accessToken - Valid access token with Channel.ReadBasic.All scope
   * @param teamId - The unique identifier of the team
   * @returns Object containing an array of channels in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getTeamChannels(token, 'team-id');
   * const channels = response.value;
   * channels.forEach(ch => console.log(ch.displayName));
   * ```
   */
  async getTeamChannels(accessToken: string, teamId: string) {
    const client = getGraphClient(accessToken);
    return await client.api(`/teams/${teamId}/channels`).get();
  },

  /**
   * Get messages from a specific channel
   *
   * Returns the most recent messages from the channel.
   *
   * @param accessToken - Valid access token with ChannelMessage.Read.All scope
   * @param teamId - The unique identifier of the team
   * @param channelId - The unique identifier of the channel
   * @returns Object containing an array of messages in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getChannelMessages(token, 'team-id', 'channel-id');
   * const messages = response.value;
   * messages.forEach(msg => console.log(msg.body.content));
   * ```
   */
  async getChannelMessages(accessToken: string, teamId: string, channelId: string) {
    const client = getGraphClient(accessToken);
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages`).get();
  },

  /**
   * Send a new message to a specific channel
   *
   * Requires the ChannelMessage.Send permission.
   *
   * @param accessToken - Valid access token with ChannelMessage.Send scope
   * @param teamId - The unique identifier of the team
   * @param channelId - The unique identifier of the channel
   * @param message - The text content of the message to send
   * @returns The created message object
   * @throws Error if the API call fails or permission is denied
   *
   * @example
   * ```typescript
   * const newMessage = await graphAPI.sendChannelMessage(
   *   token,
   *   'team-id',
   *   'channel-id',
   *   'Hello, team! 👋'
   * );
   * console.log('Message sent:', newMessage.id);
   * ```
   */
  async sendChannelMessage(
    accessToken: string,
    teamId: string,
    channelId: string,
    message: string
  ) {
    const client = getGraphClient(accessToken);
    const chatMessage = {
      body: {
        content: message,
      },
    };
    return await client.api(`/teams/${teamId}/channels/${channelId}/messages`).post(chatMessage);
  },

  /**
   * Get all members of a specific team
   *
   * @param accessToken - Valid access token with TeamMember.Read.All scope
   * @param teamId - The unique identifier of the team
   * @returns Object containing an array of team members in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getTeamMembers(token, 'team-id');
   * const members = response.value;
   * members.forEach(member => console.log(member.displayName, member.roles));
   * ```
   */
  async getTeamMembers(accessToken: string, teamId: string) {
    const client = getGraphClient(accessToken);
    return await client.api(`/teams/${teamId}/members`).get();
  },

  /**
   * Get all chats for the current user
   *
   * @param accessToken - Valid access token with Chat.Read scope
   * @returns Object containing an array of chats in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getMyChats(token);
   * const chats = response.value;
   * chats.forEach(chat => console.log(chat.topic || 'Unnamed chat'));
   * ```
   */
  async getMyChats(accessToken: string) {
    const client = getGraphClient(accessToken);
    return await client.api("/me/chats").get();
  },

  /**
   * Get messages from a specific chat
   *
   * @param accessToken - Valid access token with Chat.Read scope
   * @param chatId - The unique identifier of the chat
   * @returns Object containing an array of messages in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getChatMessages(token, 'chat-id');
   * const messages = response.value;
   * messages.forEach(msg => console.log(msg.from.user?.displayName, msg.body.content));
   * ```
   */
  async getChatMessages(accessToken: string, chatId: string) {
    const client = getGraphClient(accessToken);
    return await client.api(`/chats/${chatId}/messages`).get();
  },

  /**
   * Get calendar events for the current user
   *
   * @param accessToken - Valid access token with Calendars.Read scope
   * @returns Object containing an array of calendar events in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getMyCalendarEvents(token);
   * const events = response.value;
   * events.forEach(event => console.log(event.subject, event.start.dateTime));
   * ```
   */
  async getMyCalendarEvents(accessToken: string) {
    const client = getGraphClient(accessToken);
    return await client.api("/me/calendar/events").get();
  },

  /**
   * Get mail messages for the current user
   *
   * @param accessToken - Valid access token with Mail.Read scope
   * @param top - Number of messages to retrieve (default: 10)
   * @returns Object containing an array of mail messages in the 'value' property
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * const response = await graphAPI.getMyMail(token, 20);
   * const messages = response.value;
   * messages.forEach(msg => console.log(msg.subject, msg.from.emailAddress.name));
   * ```
   */
  async getMyMail(accessToken: string, top: number = 10) {
    const client = getGraphClient(accessToken);
    return await client.api("/me/messages").top(top).get();
  },

  /**
   * Make a custom request to any Microsoft Graph API endpoint
   *
   * Provides flexibility to call any Graph API endpoint not covered by
   * the predefined helper methods.
   *
   * @param accessToken - Valid access token with appropriate scopes
   * @param endpoint - The Graph API endpoint (e.g., '/me/presence' or 'users/{id}')
   * @param method - HTTP method to use (default: "GET")
   * @param body - Request body for POST/PATCH requests (optional)
   * @returns API response data
   * @throws Error if the API call fails
   *
   * @example
   * ```typescript
   * // Get user presence
   * const presence = await graphAPI.customRequest(token, '/me/presence', 'GET');
   * console.log('User status:', presence.availability);
   * ```
   *
   * @example
   * ```typescript
   * // Update user profile
   * await graphAPI.customRequest(
   *   token,
   *   '/me',
   *   'PATCH',
   *   { jobTitle: 'Senior Developer' }
   * );
   * ```
   */
  async customRequest(accessToken: string, endpoint: string, method: "GET" | "POST" | "PATCH" | "DELETE" = "GET", body?: unknown) {
    const client = getGraphClient(accessToken);
    const request = client.api(endpoint);

    switch (method) {
      case "GET":
        return await request.get();
      case "POST":
        return await request.post(body);
      case "PATCH":
        return await request.patch(body);
      case "DELETE":
        return await request.delete();
      default:
        return await request.get();
    }
  },
};

/**
 * Helper function to convert Graph API errors into user-friendly messages
 *
 * Translates common HTTP status codes and error messages into readable text
 * that can be displayed to users.
 *
 * @param error - The error object from a failed Graph API call
 * @returns A user-friendly error message string
 *
 * @example
 * ```typescript
 * try {
 *   await graphAPI.getMyTeams(token);
 * } catch (error) {
 *   const friendlyMessage = handleGraphError(error);
 *   alert(friendlyMessage); // "Unauthorized. Please sign in again."
 * }
 * ```
 */
export function handleGraphError(error: unknown): string {
  const err = error as { statusCode?: number; message?: string };
  if (err.statusCode === 401) {
    return "Unauthorized. Please sign in again.";
  } else if (err.statusCode === 403) {
    return "Forbidden. You don't have permission to access this resource.";
  } else if (err.statusCode === 404) {
    return "Resource not found.";
  } else if (err.message) {
    return err.message;
  } else {
    return "An error occurred while accessing Microsoft Graph API.";
  }
}
