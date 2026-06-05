import {
  DynamoDBClient,
} from "@aws-sdk/client-dynamodb";
import {
  DynamoDBDocumentClient,
  GetCommand,
  PutCommand,
} from "@aws-sdk/lib-dynamodb";
import {
  GetParameterCommand,
  PutParameterCommand,
  SSMClient,
} from "@aws-sdk/client-ssm";

import { config } from "../config";
import { SlackTokenPayload, StoredProviderCredentials, TeamsTokenPayload } from "../types";

type Provider = "slack" | "teams";

const ddbClient = new DynamoDBClient({
  region: config.awsRegion || undefined,
});

const ddb = DynamoDBDocumentClient.from(ddbClient, {
  marshallOptions: { removeUndefinedValues: true },
});

const ssm = new SSMClient({
  region: config.awsRegion || undefined,
});

function assertOrganizationId(organizationId: string): string {
  const value = String(organizationId || "").trim();
  if (!value) {
    throw new Error("organizationId is required");
  }

  if (!/^[a-zA-Z0-9:_\-/]+$/.test(value)) {
    throw new Error("organizationId contains unsupported characters");
  }

  return value;
}

function parameterNameFor(provider: Provider, organizationId: string): string {
  const normalizedPrefix = config.communicationTokensSsmPrefix.replace(/\/+$/, "");
  return `${normalizedPrefix}/${organizationId}/${provider}`;
}

interface PersistOptions {
  provider: Provider;
  organizationId: string;
  secretPayload: SlackTokenPayload | TeamsTokenPayload;
  expiresAt?: number;
  metadata?: Record<string, string | number | null>;
}

export async function persistProviderCredentials(options: PersistOptions): Promise<StoredProviderCredentials> {
  const organizationId = assertOrganizationId(options.organizationId);
  const provider = options.provider;
  const parameterName = parameterNameFor(provider, organizationId);
  const now = Date.now();

  const parameterInput: {
    Name: string;
    Type: "SecureString";
    Value: string;
    Overwrite: true;
    KeyId?: string;
  } = {
    Name: parameterName,
    Type: "SecureString",
    Value: JSON.stringify(options.secretPayload),
    Overwrite: true,
  };

  if (config.communicationTokensKmsKeyId) {
    parameterInput.KeyId = config.communicationTokensKmsKeyId;
  }

  await ssm.send(new PutParameterCommand(parameterInput));

  const record: StoredProviderCredentials = {
    provider,
    organizationId,
    parameterName,
    updatedAt: now,
    expiresAt: options.expiresAt,
    metadata: options.metadata,
  };

  await ddb.send(
    new PutCommand({
      TableName: config.communicationTokensTableName,
      Item: record,
    })
  );

  return record;
}

export async function getProviderCredentialRecord(provider: Provider, organizationId: string): Promise<StoredProviderCredentials | null> {
  const normalizedOrg = assertOrganizationId(organizationId);

  const response = await ddb.send(
    new GetCommand({
      TableName: config.communicationTokensTableName,
      Key: {
        organizationId: normalizedOrg,
        provider,
      },
    })
  );

  return (response.Item as StoredProviderCredentials | undefined) || null;
}

async function readSecurePayload<T>(parameterName: string): Promise<T> {
  const response = await ssm.send(
    new GetParameterCommand({
      Name: parameterName,
      WithDecryption: true,
    })
  );

  const raw = response.Parameter?.Value;
  if (!raw) {
    throw new Error("Missing secure credential payload");
  }

  return JSON.parse(raw) as T;
}

export async function getSlackCredentials(organizationId: string): Promise<SlackTokenPayload> {
  const record = await getProviderCredentialRecord("slack", organizationId);
  if (!record) {
    throw new Error("Slack is not connected for this organization");
  }

  return readSecurePayload<SlackTokenPayload>(record.parameterName);
}

export async function getTeamsCredentials(organizationId: string): Promise<TeamsTokenPayload> {
  const record = await getProviderCredentialRecord("teams", organizationId);
  if (!record) {
    throw new Error("Teams is not connected for this organization");
  }

  const payload = await readSecurePayload<TeamsTokenPayload>(record.parameterName);
  if (payload.expiresAt && payload.expiresAt <= Date.now()) {
    throw new Error("Stored Teams token is expired. Please reconnect Teams.");
  }

  return payload;
}
