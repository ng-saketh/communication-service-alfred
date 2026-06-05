import { SendFilePayload } from "../types";
import { requireAccessToken } from "../utils/auth";

const SLACK_API = "https://slack.com/api";
const ALLOWED_FILE_TYPES = new Set(["image/png", "application/pdf"]);

interface SlackSendInput {
  botToken: string;
  channelId: string;
  message?: string;
  threadTs?: string;
  file?: SendFilePayload;
}

interface SlackApiResponse {
  ok: boolean;
  error?: string;
  response_metadata?: {
    messages?: string[];
  };
}

function getSlackErrorMessage(data: SlackApiResponse | null, fallback: string): string {
  const details = data?.response_metadata?.messages?.join("; ");
  return details || data?.error || fallback;
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

function parseFilePayload(file?: SendFilePayload): { fileName: string; mimeType: string; content: Buffer } | null {
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
  const content = Buffer.from(normalizedBase64, "base64");
  if (!content.length) {
    throw new Error("Decoded file content is empty");
  }

  return { fileName, mimeType, content };
}

export async function listSlackChannels(botToken: string, types: "public_channel,private_channel" | "im") {
  const bearerToken = requireAccessToken(botToken, "Slack");
  const res = await fetch(
    `${SLACK_API}/conversations.list?types=${types}&exclude_archived=true&limit=200`,
    {
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        "Content-Type": "application/json",
      },
    }
  );

  const data = (await res.json()) as { ok: boolean; channels?: unknown[]; error?: string };
  if (!res.ok || !data.ok) {
    throw new Error(data.error ?? `Slack API error (${res.status})`);
  }

  return data.channels ?? [];
}

async function getUserInfo(botToken: string, userId: string): Promise<{ real_name?: string; profile?: { display_name?: string } } | null> {
  const bearerToken = requireAccessToken(botToken, "Slack");
  const res = await fetch(`${SLACK_API}/users.info?user=${userId}`, {
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
  });

  const data = (await res.json()) as { ok: boolean; user?: { real_name?: string; profile?: { display_name?: string } }; error?: string };
  if (!res.ok || !data.ok) {
    return null;
  }

  return data.user ?? null;
}

export async function listSlackDMs(botToken: string) {
  const channels = await listSlackChannels(botToken, "im");

  const enrichedChannels = await Promise.all(
    channels.map(async (channel: any) => {
      const userId = channel.user;
      if (!userId) {
        return channel;
      }

      const userInfo = await getUserInfo(botToken, userId);
      const displayName = userInfo?.profile?.display_name || userInfo?.real_name || userId;

      return {
        ...channel,
        user_name: displayName,
      };
    })
  );

  return enrichedChannels;
}

export async function sendSlackMessage(input: SlackSendInput) {
  const bearerToken = requireAccessToken(input.botToken, "Slack");
  const channelId = String(input.channelId || "").trim();
  const message = typeof input.message === "string" ? input.message : "";

  if (!channelId) {
    throw new Error("channelId is required");
  }

  const parsedFile = parseFilePayload(input.file);

  if (!message.trim() && !parsedFile) {
    throw new Error("Provide message and/or file payload");
  }

  if (parsedFile) {
    // Step 1: Request a one-time upload URL from Slack.
    const uploadUrlBody = new URLSearchParams({
      filename: parsedFile.fileName,
      length: String(parsedFile.content.length),
    });

    const uploadUrlRes = await fetch(`${SLACK_API}/files.getUploadURLExternal`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: uploadUrlBody,
    });

    const uploadUrlData = (await parseJsonSafe<
      SlackApiResponse & {
        upload_url?: string;
        file_id?: string;
      }
    >(uploadUrlRes)) || { ok: false, error: "invalid_response" };

    if (!uploadUrlRes.ok || !uploadUrlData.ok || !uploadUrlData.upload_url || !uploadUrlData.file_id) {
      throw new Error(
        getSlackErrorMessage(uploadUrlData, `Slack getUploadURLExternal failed (${uploadUrlRes.status})`)
      );
    }

    // Step 2: Upload raw bytes to the upload URL.
    const binaryUploadRes = await fetch(uploadUrlData.upload_url, {
      method: "POST",
      headers: {
        "Content-Type": parsedFile.mimeType,
      },
      body: new Uint8Array(parsedFile.content),
    });

    if (!binaryUploadRes.ok) {
      throw new Error(`Slack binary upload failed (${binaryUploadRes.status})`);
    }

    // Step 3: Complete upload and share it to a channel/thread.
    const completeUploadBody = new URLSearchParams({
      files: JSON.stringify([{ id: uploadUrlData.file_id, title: parsedFile.fileName }]),
      channel_id: channelId,
    });

    if (message.trim()) {
      completeUploadBody.set("initial_comment", message.trim());
    }

    if (input.threadTs) {
      completeUploadBody.set("thread_ts", input.threadTs);
    }

    const completeUploadRes = await fetch(`${SLACK_API}/files.completeUploadExternal`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${bearerToken}`,
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: completeUploadBody,
    });

    const completeUploadData = (await parseJsonSafe<
      SlackApiResponse & {
        files?: Array<{ id?: string; name?: string; mimetype?: string; permalink?: string }>;
      }
    >(completeUploadRes)) || { ok: false, error: "invalid_response" };

    if (!completeUploadRes.ok || !completeUploadData.ok) {
      throw new Error(
        getSlackErrorMessage(completeUploadData, `Slack completeUploadExternal failed (${completeUploadRes.status})`)
      );
    }

    const uploadedFile = completeUploadData.files?.[0];

    return {
      channelId,
      file: {
        id: uploadedFile?.id,
        name: uploadedFile?.name || parsedFile.fileName,
        mimetype: uploadedFile?.mimetype || parsedFile.mimeType,
        permalink: uploadedFile?.permalink,
      },
    };
  }

  const res = await fetch(`${SLACK_API}/chat.postMessage`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${bearerToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      channel: channelId,
      text: message,
      thread_ts: input.threadTs,
    }),
  });

  const data = (await parseJsonSafe<SlackApiResponse & { channel?: string; ts?: string }>(res)) || { ok: false };
  if (!res.ok || !data.ok) {
    throw new Error(getSlackErrorMessage(data, `Slack API error (${res.status})`));
  }

  return {
    channelId,
    channel: data.channel || channelId,
    ts: data.ts,
  };
}
