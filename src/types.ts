export interface SlackTokenPayload {
  botToken: string;
  userToken: string | null;
  teamId: string | null;
  teamName: string | null;
  botUserId: string | null;
  scope: string | null;
  userScope: string | null;
  appId: string | null;
  authedUserId: string | null;
  expiresAt: number;
}

export interface TeamsTokenPayload {
  accessToken: string;
  expiresAt: number;
}

export interface StoredProviderCredentials {
  provider: "slack" | "teams";
  organizationId: string;
  parameterName: string;
  updatedAt: number;
  expiresAt?: number;
  metadata?: Record<string, string | number | null>;
}

export interface SendFilePayload {
  fileName: string;
  mimeType: string;
  contentBase64: string;
}
