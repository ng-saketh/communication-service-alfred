/**
 * MSAL (Microsoft Authentication Library) client for Teams OAuth 2.0
 * Handles automatic PKCE and code exchange.
 */

import { ConfidentialClientApplication, LogLevel } from "@azure/msal-node";
import { CryptoProvider } from "@azure/msal-node";
import { config } from "../config";

// Initialize MSAL client with Teams credentials
const msalConfig = {
  auth: {
    clientId: config.teamsClientId,
    clientSecret: config.teamsClientSecret,
    authority: "https://login.microsoftonline.com/common",
  },
  system: {
    loggerOptions: {
      loggerCallback(_logLevel: LogLevel, _message: string) {
        // Keep callback to satisfy MSAL typing while avoiding noisy logs.
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Error,
    },
  },
};

const msalClient = new ConfidentialClientApplication(msalConfig);
const cryptoProvider = new CryptoProvider();

interface PkceCodes {
  verifier: string;
  challenge: string;
}

export async function generatePkceCodes(): Promise<PkceCodes> {
  const codes = await cryptoProvider.generatePkceCodes();
  return {
    verifier: codes.verifier,
    challenge: codes.challenge,
  };
}

/**
 * Generate Microsoft OAuth authorization URL with PKCE
 * @param state - CSRF protection token (typically signed JWT)
 * @param codeChallenge - PKCE S256 code challenge
 * @returns Authorization URL for redirect to Microsoft login
 */
export async function getTeamsAuthUrl(state: string, codeChallenge: string): Promise<string> {
  const authCodeUrlParameters = {
    scopes: config.teamsScopes,
    redirectUri: config.teamsRedirectUri,
    state, // MSAL includes state parameter for CSRF protection
    codeChallenge,
    codeChallengeMethod: "S256",
  };

  const authUrl = await msalClient.getAuthCodeUrl(authCodeUrlParameters);
  return authUrl;
}

/**
 * Exchange authorization code for tokens
 * @param code - Authorization code from OAuth callback
 * @param redirectUri - Must match registered redirect URI
 * @param codeVerifier - PKCE code verifier used to generate the code challenge
 * @returns Tokens: access_token, expires_in
 */
export async function exchangeCodeForTokens(code: string, redirectUri: string, codeVerifier: string) {
  const tokenRequest = {
    code,
    redirectUri,
    scopes: config.teamsScopes,
    codeVerifier,
  };

  const response = await msalClient.acquireTokenByCode(tokenRequest);

  return {
    accessToken: response?.accessToken || "",
    expiresIn: response?.expiresOn
      ? Math.floor((response.expiresOn.getTime() - Date.now()) / 1000)
      : 3600,
  };
}
