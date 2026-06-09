import { SendFilePayload } from "../types";
import { requireAccessToken } from "../utils/auth";

const GRAPH_API = "https://graph.microsoft.com/v1.0";
const ALLOWED_FILE_TYPES = new Set(["image/png", "application/pdf"]);

interface TeamsSendInput {
  accessToken: string;
  destinationType: "chat" | "channel" | "user";
  destinationId: string;
  teamId?: string;
  message?: string;
  file?: SendFilePayload;
}

interface GraphApiResponse {
  error?: {
    code?: string;
    message?: string;
  };
}

interface GraphCollectionResponse<T> {
  value?: T[];
  "@odata.nextLink"?: string;
  error?: {
    code?: string;
    message?: string;
  };
}

interface TeamsChatMember {
  id?: string;
  displayName?: string;
  userId?: string;
  email?: string;
}

interface TeamsDirectoryUser {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

interface TeamsChatItem {
  id?: string;
  topic?: string | null;
  chatType?: string;
  [key: string]: unknown;
}

function getGraphErrorMessage(data: GraphApiResponse | null, fallback: string): string {
  return data?.error?.message || fallback;
}

function isInsufficientPrivilegesError(message: string): boolean {
  return /insufficient privileges/i.test(message);
}

function encodeGraphPathSegment(value: string): string {
  return encodeURIComponent(value);
}

async function parseJsonSafe<T>(response: Response): Promise<T | null> {
  const text = await response.text();
  if (!text) {
    return null;
  }

  try {
    return JSON.parse(text) as T;
  } catch {
    return null;
  }
}

async function fetchGraphCollection<T>(accessToken: string, url: string, maxItems?: number): Promise<T[]> {
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const items: T[] = [];
  let nextUrl: string | null = url;

  while (nextUrl) {
    const currentUrl = nextUrl;
    const res: Response = await fetch(currentUrl, {
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        "Content-Type": "application/json",
      },
    });

    const data: GraphCollectionResponse<T> | null = await parseJsonSafe<GraphCollectionResponse<T>>(res);
    if (!res.ok) {
      const message = data?.error?.message || `Teams API error (${res.status})`;
      throw new Error(message);
    }

    const pageItems = data?.value ?? [];
    if (typeof maxItems === "number") {
      const remaining = maxItems - items.length;
      if (remaining <= 0) {
        break;
      }

      items.push(...pageItems.slice(0, remaining));
      if (items.length >= maxItems) {
        break;
      }
    } else {
      items.push(...pageItems);
    }

    nextUrl = data?.["@odata.nextLink"] || null;
  }

  return items;
}

async function getCurrentUser(accessToken: string): Promise<{ id: string; displayName: string } | null> {
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const res = await fetch(`${GRAPH_API}/me?$select=id,displayName`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!res.ok) {
    return null;
  }

  const data = await parseJsonSafe<{ id?: string; displayName?: string }>(res);
  if (!data?.id) {
    return null;
  }

  return {
    id: data.id,
    displayName: data.displayName || "",
  };
}

async function getDirectoryUserById(accessToken: string, userId: string): Promise<TeamsDirectoryUser | null> {
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const res = await fetch(`${GRAPH_API}/users/${encodeGraphPathSegment(userId)}?$select=id,displayName,mail,userPrincipalName`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!res.ok) {
    return null;
  }

  return parseJsonSafe<TeamsDirectoryUser>(res);
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

async function resolveOrCreateOneOnOneChatId(accessToken: string, targetUserId: string): Promise<string> {
  const [currentUser, chats] = await Promise.all([
    getCurrentUser(accessToken),
    fetchGraphCollection<TeamsChatItem>(accessToken, `${GRAPH_API}/me/chats?$top=50`),
  ]);

  const oneOnOneChats = chats.filter((chat) => chat.chatType === "oneOnOne" && chat.id);
  for (const chat of oneOnOneChats) {
    const members = await getTeamsChatMembers(accessToken, String(chat.id));
    if (members.some((member) => member.userId === targetUserId)) {
      return String(chat.id);
    }
  }

  if (!currentUser?.id) {
    throw new Error("Unable to resolve current Teams user profile to create direct chat");
  }

  const bearerToken = requireAccessToken(accessToken, "Teams");
  const createPayload = {
    chatType: "oneOnOne",
    members: [
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${escapeODataString(currentUser.id)}')`,
      },
      {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        roles: ["owner"],
        "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${escapeODataString(targetUserId)}')`,
      },
    ],
  };

  const createRes = await fetch(`${GRAPH_API}/chats`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(createPayload),
  });

  const createData = await parseJsonSafe<GraphApiResponse & { id?: string }>(createRes);
  if (!createRes.ok || createData?.error || !createData?.id) {
    const message = getGraphErrorMessage(createData, `Teams API error (${createRes.status})`);
    if (isInsufficientPrivilegesError(message)) {
      throw new Error(
        "Insufficient privileges to create one-on-one chat. Reconnect Teams with delegated scopes including Chat.Create, Chat.ReadWrite, ChatMessage.Send, User.Read, and User.ReadBasic.All."
      );
    }
    throw new Error(`Unable to create one-on-one chat for teammate: ${message}`);
  }

  return createData.id;
}

function parseFilePayload(file?: SendFilePayload): { fileName: string; mimeType: string; contentBase64: string } | null {
  if (!file) {
    return null;
  }

  const fileName = String(file.fileName || "").trim();
  const mimeType = String(file.mimeType || "").trim();
  const contentBase64 = String(file.contentBase64 || "").trim();

  if (!fileName) {
    throw new Error("Missing file.fileName");
  }
  if (!contentBase64) {
    throw new Error("Missing file.contentBase64");
  }
  if (!ALLOWED_FILE_TYPES.has(mimeType)) {
    throw new Error("Only PNG and PDF files are supported");
  }

  const normalizedBase64 = contentBase64.replace(/\s/g, "");
  const decoded = Buffer.from(normalizedBase64, "base64");
  if (!decoded.length) {
    throw new Error("Decoded file content is empty");
  }

  return {
    fileName,
    mimeType,
    contentBase64: normalizedBase64,
  };
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildTeamsMessageBody(message: string, parsedFile: { fileName: string; mimeType: string; contentBase64: string } | null): {
  body: {
    contentType: "html";
    content: string;
  };
  hostedContents?: Array<{
    "@microsoft.graph.temporaryId": string;
    contentBytes: string;
    contentType: string;
  }>;
} {
  const normalizedMessage = message.trim();

  if (!parsedFile) {
    return {
      body: {
        contentType: "html",
        content: normalizedMessage || "&nbsp;",
      },
    };
  }

  // Teams can render inline images through hostedContents.
  if (parsedFile.mimeType === "image/png") {
    const contentParts: string[] = [];
    if (normalizedMessage) {
      contentParts.push(escapeHtml(normalizedMessage));
    }
    contentParts.push(`<img alt="${escapeHtml(parsedFile.fileName)}" src="../hostedContents/1/$value" />`);

    return {
      body: {
        contentType: "html",
        content: contentParts.join("<br/><br/>"),
      },
      hostedContents: [
        {
          "@microsoft.graph.temporaryId": "1",
          contentBytes: parsedFile.contentBase64,
          contentType: parsedFile.mimeType,
        },
      ],
    };
  }

  const pdfNotice = `<strong>File attachment:</strong> ${escapeHtml(parsedFile.fileName)}`;
  return {
    body: {
      contentType: "html",
      content: normalizedMessage ? `${escapeHtml(normalizedMessage)}<br/><br/>${pdfNotice}` : pdfNotice,
    },
  };
}

async function getTeamsChatMembers(accessToken: string, chatId: string) {
  try {
    return await fetchGraphCollection<TeamsChatMember>(accessToken, `${GRAPH_API}/me/chats/${encodeGraphPathSegment(chatId)}/members?$top=50`);
  } catch {
    return [];
  }
}

export async function getTeamsChats(accessToken: string) {
  const [chats, currentUser] = await Promise.all([
    fetchGraphCollection<TeamsChatItem>(accessToken, `${GRAPH_API}/me/chats?$top=50`),
    getCurrentUser(accessToken),
  ]);

  const userLookupCache = new Map<string, string>();

  // Fetch members for each chat to get colleague names for one-on-one chats
  const chatsWithMembers = await Promise.all(
    chats.map(async (chat) => {
      if (!chat.id) return chat;

      try {
        const members = await getTeamsChatMembers(accessToken, chat.id);
        
        // For one-on-one chats (topic is null), use teammate displayName (exclude current user)
        let displayName = chat.topic;
        if (!displayName && members.length > 0) {
          const colleague = members.find((m) => {
            if (!m.displayName) {
              return false;
            }

            if (!currentUser) {
              return true;
            }

            if (m.userId && m.userId === currentUser.id) {
              return false;
            }

            if (!m.userId && currentUser.displayName && m.displayName === currentUser.displayName) {
              return false;
            }

            return true;
          });
          if (colleague) {
            displayName = colleague.displayName;
          } else {
            displayName = members
              .map((member) => member.displayName)
              .filter((name): name is string => Boolean(name && (!currentUser || name !== currentUser.displayName)))
              .join(", ");

            if (!displayName) {
              const colleagueMember = members.find((member) => member.userId && (!currentUser || member.userId !== currentUser.id));
              if (colleagueMember?.userId) {
                const cached = userLookupCache.get(colleagueMember.userId);
                if (cached) {
                  displayName = cached;
                } else {
                  const profile = await getDirectoryUserById(accessToken, colleagueMember.userId);
                  const resolvedName = (profile?.displayName || profile?.mail || profile?.userPrincipalName || "").trim();
                  if (resolvedName) {
                    userLookupCache.set(colleagueMember.userId, resolvedName);
                    displayName = resolvedName;
                  }
                }
              }
            }
          }
        }

        return {
          ...chat,
          displayName,
          members,
        };
      } catch {
        // If member fetch fails, return chat with original data
        return chat;
      }
    })
  );

  return chatsWithMembers;
}

export async function getTeamsPeople(accessToken: string) {
  let users: TeamsDirectoryUser[] = [];
  let currentUser: { id: string; displayName: string } | null = null;

  try {
    [users, currentUser] = await Promise.all([
      fetchGraphCollection<TeamsDirectoryUser>(
        accessToken,
        `${GRAPH_API}/users?$select=id,displayName,mail,userPrincipalName&$top=100`,
        300
      ),
      getCurrentUser(accessToken),
    ]);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to list Teams people";
    if (isInsufficientPrivilegesError(message)) {
      console.warn("Insufficient privileges to list Teams teammates. User.Read/User.ReadBasic.All scopes may be missing. Teammates section will be empty.");
      return [];
    }

    console.error("Error listing Teams people:", message);
    return [];
  }

  return users
    .filter((user) => Boolean(user.id) && user.id !== currentUser?.id)
    .map((user) => {
      const displayName = (user.displayName || user.mail || user.userPrincipalName || "").trim();
      return {
        id: String(user.id),
        displayName,
        email: user.mail || user.userPrincipalName || "",
      };
    })
    .filter((user) => Boolean(user.displayName))
    .sort((a, b) => a.displayName.localeCompare(b.displayName));
}

export async function getTeamsChannels(accessToken: string) {
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const res = await fetch(`${GRAPH_API}/me/joinedTeams`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  const data = (await res.json()) as {
    value?: Array<{ id?: string; displayName?: string }>;
    error?: unknown;
  };
  if (!res.ok) {
    throw new Error(`Teams API error (${res.status})`);
  }

  const teams = data.value ?? [];
  const channels = [];

  for (const team of teams) {
    if (!team.id) continue;

    const channelsRes = await fetch(`${GRAPH_API}/teams/${team.id}/channels`, {
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        "Content-Type": "application/json",
      },
    });

    const channelsData = (await channelsRes.json()) as {
      value?: Array<{ id?: string; displayName?: string }>;
      error?: unknown;
    };

    if (channelsRes.ok && channelsData.value) {
      for (const channel of channelsData.value) {
        channels.push({
          id: channel.id,
          displayName: channel.displayName,
          teamId: team.id,
          teamDisplayName: team.displayName,
        });
      }
    }
  }

  return channels;
}

export async function getTeamsGroups(accessToken: string) {
  const [chats, channels] = await Promise.all([
    getTeamsChats(accessToken),
    getTeamsChannels(accessToken),
  ]);

  // Filter chats to only non-oneOnOne (group and meeting chats)
  const groupChats = chats
    .filter((chat) => chat.chatType !== "oneOnOne" && chat.id)
    .map((chat) => ({
      id: chat.id,
      displayName: chat.displayName || chat.topic || "Unnamed Group",
      type: "chat",
    }));

  // Rename channels to include teamDisplayName for context
  const formattedChannels = channels.map((channel) => ({
    id: channel.id,
    displayName: channel.teamDisplayName ? `${channel.teamDisplayName} - ${channel.displayName}` : channel.displayName,
    type: "channel",
    teamId: channel.teamId,
  }));

  // Combine group chats and channels, sorted by displayName
  const groups = [...groupChats, ...formattedChannels].sort((a, b) =>
    (a.displayName || "").localeCompare(b.displayName || "")
  );

  return groups;
}

export async function sendTeamsMessage(input: TeamsSendInput) {
  const bearerToken = requireAccessToken(input.accessToken, "Teams");
  const destinationId = String(input.destinationId || "").trim();
  const message = typeof input.message === "string" ? input.message : "";

  if (!destinationId) {
    throw new Error("destinationId is required");
  }

  const parsedFile = parseFilePayload(input.file);

  if (!message.trim() && !parsedFile) {
    throw new Error("Provide message and/or file payload");
  }

  let endpoint = "";
  let resolvedDestinationId = destinationId;
  if (input.destinationType === "chat") {
    endpoint = `${GRAPH_API}/chats/${encodeGraphPathSegment(resolvedDestinationId)}/messages`;
  } else if (input.destinationType === "channel" && input.teamId) {
    endpoint = `${GRAPH_API}/teams/${encodeGraphPathSegment(input.teamId)}/channels/${encodeGraphPathSegment(resolvedDestinationId)}/messages`;
  } else if (input.destinationType === "user") {
    resolvedDestinationId = await resolveOrCreateOneOnOneChatId(input.accessToken, destinationId);
    endpoint = `${GRAPH_API}/chats/${encodeGraphPathSegment(resolvedDestinationId)}/messages`;
  } else {
    throw new Error("Invalid destination type or missing teamId for channel");
  }

  const messageBody = buildTeamsMessageBody(message, parsedFile);

  const res = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(messageBody),
  });

  const data = (await parseJsonSafe<GraphApiResponse & { id?: string }>(res)) || { error: { message: "invalid_response" } };

  if (!res.ok || data.error) {
    const graphMessage = getGraphErrorMessage(data, `Teams API error (${res.status})`);
    if (graphMessage === "NotFound") {
      throw new Error(
        `NotFound: Teams destination '${destinationId}' was not found. Ensure the full destination id from /teams/chats or /teams/channels is sent (not a truncated value like '19').`
      );
    }

    throw new Error(graphMessage);
  }

  return {
    destinationType: input.destinationType,
    destinationId,
    messageId: data.id || "",
  };
}
