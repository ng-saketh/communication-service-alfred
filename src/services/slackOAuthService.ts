import axios from "axios";
import { config } from "../config";
import { SlackTokenPayload } from "../types";

interface SlackOAuthResponse {
  ok: boolean;
  error?: string;
  access_token?: string;
  scope?: string;
  app_id?: string;
  team?: {
    id?: string;
    name?: string;
  };
  bot_user_id?: string;
  authed_user?: {
    id?: string;
    scope?: string;
    access_token?: string;
  };
}

export async function exchangeCodeForTokens(code: string): Promise<SlackTokenPayload> {
  const body = new URLSearchParams({
    client_id: config.slackClientId,
    client_secret: config.slackClientSecret,
    code,
    redirect_uri: config.slackRedirectUri,
  }).toString();

  const response = await axios.post<SlackOAuthResponse>("https://slack.com/api/oauth.v2.access", body, {
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    timeout: 15000,
  });

  const data = response.data;
  if (!data.ok || !data.access_token) {
    throw new Error(data.error || "Slack OAuth token exchange failed");
  }

  return {
    botToken: data.access_token,
    userToken: data.authed_user?.access_token || null,
    teamId: data.team?.id || null,
    teamName: data.team?.name || null,
    botUserId: data.bot_user_id || null,
    scope: data.scope || null,
    userScope: data.authed_user?.scope || null,
    appId: data.app_id || null,
    authedUserId: data.authed_user?.id || null,
    expiresAt: Date.now() + config.tokenTtlMs,
  };
}
