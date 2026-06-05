export function stripBearerPrefix(rawToken: string | undefined): string {
  return String(rawToken || "").replace(/^Bearer\s+/i, "").trim();
}

export function requireAccessToken(rawToken: string | undefined, provider: "Slack" | "Teams"): string {
  const token = stripBearerPrefix(rawToken);
  if (!token) {
    throw new Error(`Missing ${provider} access token`);
  }

  return token;
}
