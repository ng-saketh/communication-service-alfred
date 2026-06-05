import crypto from "crypto";
import { Router } from "express";

import { config } from "../config";
import { signState, verifyState } from "../state";
import { getProviderCredentialRecord, persistProviderCredentials } from "../services/credentialStore";
import { exchangeCodeForTokens } from "../services/slackOAuthService";

export const slackOAuthRouter = Router();

slackOAuthRouter.get("/install-url", (req, res) => {
  try {
    const clientContext = typeof req.query.clientContext === "string" ? req.query.clientContext : undefined;
    const organizationId = typeof req.query.organizationId === "string" ? req.query.organizationId.trim() : "";
    if (!organizationId) {
      throw new Error("organizationId is required");
    }

    const state = signState(
      {
        nonce: crypto.randomUUID(),
        requestedAt: Date.now(),
        clientContext,
        organizationId,
      },
      config.stateSecret
    );

    const params = new URLSearchParams({
      client_id: config.slackClientId,
      scope: config.slackScopes,
      redirect_uri: config.slackRedirectUri,
      state,
    });

    if (config.slackUserScopes) {
      params.set("user_scope", config.slackUserScopes);
    }

    return res.status(200).json({
      installUrl: `${config.slackBaseUrl}?${params.toString()}`,
      state,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to generate install URL";
    return res.status(400).json({ status: "error", message });
  }
});

slackOAuthRouter.get("/callback", async (req, res) => {
  try {
    const code = typeof req.query.code === "string" ? req.query.code : "";
    const state = typeof req.query.state === "string" ? req.query.state : "";

    if (!code || !state) {
      throw new Error("Missing code or state");
    }

    const verifiedState = verifyState(state, config.stateSecret);
    if (!verifiedState.organizationId) {
      throw new Error("Missing organizationId in OAuth state");
    }

    const tokenPayload = await exchangeCodeForTokens(code);

    await persistProviderCredentials({
      provider: "slack",
      organizationId: verifiedState.organizationId,
      secretPayload: tokenPayload,
      expiresAt: tokenPayload.expiresAt,
      metadata: {
        teamId: tokenPayload.teamId,
        teamName: tokenPayload.teamName,
        appId: tokenPayload.appId,
        botUserId: tokenPayload.botUserId,
      },
    });

    if (config.frontendCallbackUrl) {
      const callbackUrl = new URL(config.frontendCallbackUrl);
      callbackUrl.searchParams.set("status", "success");
      callbackUrl.searchParams.set("teamId", tokenPayload.teamId || "");
      callbackUrl.searchParams.set("teamName", tokenPayload.teamName || "");
      if (verifiedState.clientContext) {
        callbackUrl.searchParams.set("clientContext", verifiedState.clientContext);
      }
      return res.redirect(302, callbackUrl.toString());
    }

    return res.status(200).json({
      status: "success",
      connected: true,
      teamId: tokenPayload.teamId,
      teamName: tokenPayload.teamName,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Slack OAuth callback failed";

    if (config.frontendCallbackUrl) {
      const callbackUrl = new URL(config.frontendCallbackUrl);
      callbackUrl.searchParams.set("status", "error");
      callbackUrl.searchParams.set("message", message);
      return res.redirect(302, callbackUrl.toString());
    }

    return res.status(400).json({ status: "error", message });
  }
});

slackOAuthRouter.get("/status", async (req, res) => {
  try {
    const organizationId = String(req.query.organizationId || "").trim();
    if (!organizationId) {
      return res.status(400).json({ message: "organizationId is required" });
    }

    const record = await getProviderCredentialRecord("slack", organizationId);
    return res.status(200).json({
      connected: Boolean(record),
      metadata: record?.metadata || null,
      updatedAt: record?.updatedAt || null,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to fetch Slack status";
    return res.status(400).json({ status: "error", message });
  }
});
