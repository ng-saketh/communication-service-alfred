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

function parseFilePayload(file?: SendFilePayload): { fileName: string } | null {
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

  return { fileName };
}

export async function getTeamsChats(accessToken: string) {
  const bearerToken = requireAccessToken(accessToken, "Teams");
  const res = await fetch(`${GRAPH_API}/me/chats`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  const data = (await res.json()) as { ok?: boolean; value?: unknown[]; error?: unknown };
  if (!res.ok) {
    throw new Error(`Teams API error (${res.status})`);
  }

  return data.value ?? [];
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
    endpoint = `${GRAPH_API}/chats/${destinationId}/messages`;
  } else if (input.destinationType === "channel" && input.teamId) {
    endpoint = `${GRAPH_API}/teams/${input.teamId}/channels/${destinationId}/messages`;
  } else {
    throw new Error("Invalid destination type or missing teamId for channel");
  }

  // Build message body
  const messageBody: {
    body: {
      contentType: "html";
      content: string;
    };
  } = {
    body: {
      contentType: "html",
      content: message || "&nbsp;",
    },
  };

  // If we have a file, we'll need to handle it differently
  if (parsedFile) {
    // For Teams, we need to create the message first, then add the attachment
    // This is a simplified approach - a full implementation would handle inline attachments
    messageBody.body.content = `${message}<br/><br/><strong>File attachment:</strong> ${parsedFile.fileName}`;
  }

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
    throw new Error(getGraphErrorMessage(data, `Teams API error (${res.status})`));
  }

  return {
    destinationType: input.destinationType,
    destinationId,
    messageId: data.id || "",
  };
}
