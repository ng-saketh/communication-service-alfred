import dotenv from "dotenv";

dotenv.config();

function getRequired(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return value;
}

function getRequiredCsv(name: string): string[] {
  const value = getRequired(name);
  const items = value
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  if (!items.length) {
    throw new Error(`Environment variable ${name} must include at least one value`);
  }

  return items;
}

function getOptionalPositiveNumber(name: string, defaultValue: number): number {
  const raw = process.env[name]?.trim();
  if (!raw) {
    return defaultValue;
  }

  const parsed = Number(raw);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`Environment variable ${name} must be a positive number`);
  }

  return parsed;
}

function getOptional(name: string, defaultValue = ""): string {
  return process.env[name]?.trim() || defaultValue;
}

export const config = {
  port: getOptionalPositiveNumber("PORT", 4012),
  frontendOrigin: process.env.FRONTEND_ORIGIN?.trim() || "*",
  frontendCallbackUrl: process.env.FRONTEND_CALLBACK_URL?.trim() || "",

  sendgridApiKey: getRequired("SENDGRID_API_KEY"),
  fromEmail: getRequired("FROM_EMAIL"),
  fromName: process.env.FROM_NAME?.trim() || "Alfred Platform",

  slackClientId: getRequired("SLACK_CLIENT_ID"),
  slackClientSecret: getRequired("SLACK_CLIENT_SECRET"),
  slackRedirectUri: getRequired("SLACK_REDIRECT_URI"),
  slackScopes: getRequired("SLACK_SCOPES"),
  slackUserScopes: process.env.SLACK_USER_SCOPES?.trim() || "",
  slackBaseUrl: process.env.SLACK_BASE_URL?.trim() || "https://slack.com/oauth/v2/authorize",

  teamsClientId: getRequired("TEAMS_CLIENT_ID"),
  teamsClientSecret: getRequired("TEAMS_CLIENT_SECRET"),
  teamsRedirectUri: getRequired("TEAMS_REDIRECT_URI"),
  teamsScopes: getRequiredCsv("TEAMS_SCOPES"),

  stateSecret: getRequired("STATE_SECRET"),
  tokenTtlMs: getOptionalPositiveNumber("TOKEN_TTL_MS", 10 * 60 * 1000),

  awsRegion: getOptional("AWS_REGION", getOptional("REGION", "")),
  communicationTokensTableName: getRequired("COMMUNICATION_TOKENS_TABLE_NAME"),
  communicationTokensSsmPrefix: getOptional("COMMUNICATION_TOKENS_SSM_PREFIX", "/alfred/communication-tokens"),
  communicationTokensKmsKeyId: getOptional("COMMUNICATION_TOKENS_KMS_KEY_ID"),
};
