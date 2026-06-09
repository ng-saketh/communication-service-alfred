import { Router } from "express";
import crypto from "crypto";
import { config } from "../config";
import { exchangeCodeForTokens, generatePkceCodes, getTeamsAuthUrl } from "../services/msalClient";
import { getProviderCredentialRecord, persistProviderCredentials, disconnectProvider } from "../services/credentialStore";
import { signState, verifyState } from "../state";

export const teamsOAuthRouter = Router();

interface PkceRecord {
  verifier: string;
  expiresAt: number;
}

const pkceStore = new Map<string, PkceRecord>();

function cleanupExpiredPkce(now: number): void {
  for (const [nonce, record] of pkceStore.entries()) {
    if (record.expiresAt <= now) {
      pkceStore.delete(nonce);
    }
  }
}

function resolveClientRedirect(clientContext?: string): string {
  const fallbackPath = "/onboarding";
  const baseOrigin = config.frontendOrigin && config.frontendOrigin !== "*" ? config.frontendOrigin : "http://localhost:5173";

  if (!clientContext) {
    return `${baseOrigin}${fallbackPath}`;
  }

  if (clientContext.startsWith("/")) {
    return `${baseOrigin}${clientContext}`;
  }

  return `${baseOrigin}${fallbackPath}`;
}

teamsOAuthRouter.get("/auth-url", async (req, res) => {
  try {
    cleanupExpiredPkce(Date.now());

    const clientContext = typeof req.query.clientContext === "string" ? req.query.clientContext : undefined;
    const organizationId = typeof req.query.organizationId === "string" ? req.query.organizationId.trim() : "";
    if (!organizationId) {
      throw new Error("organizationId is required");
    }

    const statePayload: { nonce: string; requestedAt: number; clientContext?: string; organizationId: string } = {
      nonce: crypto.randomUUID(),
      requestedAt: Date.now(),
      organizationId,
    };

    if (typeof clientContext === "string" && clientContext.length > 0) {
      statePayload.clientContext = clientContext;
    }

    const state = signState(
      statePayload,
      config.stateSecret
    );

    const pkceCodes = await generatePkceCodes();
    const authUrl = await getTeamsAuthUrl(state, pkceCodes.challenge);

    pkceStore.set(statePayload.nonce, {
      verifier: pkceCodes.verifier,
      expiresAt: Date.now() + config.tokenTtlMs,
    });

    return res.status(200).json({ authUrl });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to generate Teams auth URL";
    return res.status(400).json({ status: "error", message });
  }
});

// Handle Azure OAuth redirect - called by Azure with authorization code
teamsOAuthRouter.get("/callback", async (req, res) => {
  try {
    cleanupExpiredPkce(Date.now());

    const code = String(req.query.code || "").trim();
    const error = typeof req.query.error === "string" ? req.query.error : "";
    const errorDescription = typeof req.query.error_description === "string" ? req.query.error_description : "";
    const rawState = typeof req.query.state === "string" ? req.query.state : "";

    const statePayload = verifyState(rawState, config.stateSecret);
    if (!statePayload.organizationId) {
      throw new Error("Missing organizationId in OAuth state");
    }

    const pkceRecord = pkceStore.get(statePayload.nonce);
    if (!pkceRecord) {
      throw new Error("Missing or expired PKCE verifier for Teams OAuth flow");
    }
    pkceStore.delete(statePayload.nonce);

    const redirectBase = resolveClientRedirect(statePayload.clientContext);

    // If there was an error from Azure, redirect to frontend with error
    if (error || errorDescription) {
      const params = new URLSearchParams();
      params.set("teamsStatus", "error");
      params.set("teamsError", errorDescription || error || "OAuth failed");
      return res.redirect(`${redirectBase}${redirectBase.includes("?") ? "&" : "?"}${params.toString()}`);
    }

    if (!code) {
      const params = new URLSearchParams();
      params.set("teamsStatus", "error");
      params.set("teamsError", "Missing authorization code");
      return res.redirect(`${redirectBase}${redirectBase.includes("?") ? "&" : "?"}${params.toString()}`);
    }

    // Exchange code for tokens with PKCE verifier generated during auth URL creation.
    const tokenResult = await exchangeCodeForTokens(code, config.teamsRedirectUri, pkceRecord.verifier);
    const expiresAt = Date.now() + tokenResult.expiresIn * 1000;

    await persistProviderCredentials({
      provider: "teams",
      organizationId: statePayload.organizationId,
      secretPayload: {
        accessToken: tokenResult.accessToken,
        expiresAt,
      },
      expiresAt,
    });

    // Redirect back to frontend with non-sensitive status only
    const params = new URLSearchParams();
    params.set("teamsStatus", "success");
    params.set("teamsExpiresAt", String(expiresAt));

    return res.redirect(`${redirectBase}${redirectBase.includes("?") ? "&" : "?"}${params.toString()}`);
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to exchange Teams code for tokens";
    console.error("Teams OAuth callback error:", message);
    const baseOrigin = config.frontendOrigin && config.frontendOrigin !== "*" ? config.frontendOrigin : "http://localhost:5173";
    const params = new URLSearchParams();
    params.set("teamsStatus", "error");
    params.set("teamsError", message);
    return res.redirect(`${baseOrigin}/onboarding?${params.toString()}`);
  }
});

teamsOAuthRouter.get("/status", async (req, res) => {
  try {
    const organizationId = String(req.query.organizationId || "").trim();
    if (!organizationId) {
      return res.status(400).json({ message: "organizationId is required" });
    }

    const record = await getProviderCredentialRecord("teams", organizationId);
    return res.status(200).json({
      connected: Boolean(record),
      expiresAt: record?.expiresAt || null,
      updatedAt: record?.updatedAt || null,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to fetch Teams status";
    return res.status(400).json({ status: "error", message });
  }
});

teamsOAuthRouter.post("/disconnect", async (req, res) => {
  try {
    const organizationId = String(req.body.organizationId || "").trim();
    if (!organizationId) {
      return res.status(400).json({ status: "error", message: "organizationId is required" });
    }

    await disconnectProvider("teams", organizationId);
    return res.status(200).json({ status: "success", message: "Teams disconnected successfully" });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Failed to disconnect Teams";
    console.error("Teams disconnect error:", message);
    return res.status(400).json({ status: "error", message });
  }
});
