import { SendFilePayload } from "../types";
import { requireAccessToken } from "../utils/auth";

const GRAPH_API = "https://graph.microsoft.com/v1.0";
const ALLOWED_FILE_TYPES = new Set(["image/png", "application/pdf"]);

interface TeamsSendInput {
  accessToken: string;
  destinationType: "chat" | "channel";
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

function getGraphErrorMessage(data: GraphApiResponse | null, fallback: string): string {
  return data?.error?.message || fallback;
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
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const res = await fetch(`${GRAPH_API}/me/chats/${chatId}/members`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  if (!res.ok) {
    return [];
  }

  const data = (await res.json()) as { value?: Array<{ displayName?: string; id?: string }> };
  return data.value ?? [];
}

export async function getTeamsChats(accessToken: string) {
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const res = await fetch(`${GRAPH_API}/me/chats`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  const data = (await res.json()) as { ok?: boolean; value?: Array<{ id?: string; topic?: string | null; chatType?: string; [key: string]: unknown }>; error?: unknown };
  if (!res.ok) {
    throw new Error(`Teams API error (${res.status})`);
  }

  const chats = data.value ?? [];

  // Fetch members for each chat to get colleague names for one-on-one chats
  const chatsWithMembers = await Promise.all(
    chats.map(async (chat) => {
      if (!chat.id) return chat;

      try {
        const members = await getTeamsChatMembers(accessToken, chat.id);
        
        // For one-on-one chats (topic is null), use the colleague's displayName
        let displayName = chat.topic;
        if (!displayName && members.length > 0) {
          // Find the first member who is not the current user (excluding self)
          const colleague = members.find((m) => m.displayName);
          if (colleague) {
            displayName = colleague.displayName;
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
  if (input.destinationType === "chat") {
    endpoint = `${GRAPH_API}/chats/${encodeGraphPathSegment(destinationId)}/messages`;
  } else if (input.destinationType === "channel" && input.teamId) {
    endpoint = `${GRAPH_API}/teams/${encodeGraphPathSegment(input.teamId)}/channels/${encodeGraphPathSegment(destinationId)}/messages`;
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
