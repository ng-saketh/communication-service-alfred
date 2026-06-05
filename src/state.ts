import crypto from "crypto";

interface OAuthStatePayload {
  nonce: string;
  requestedAt: number;
  clientContext?: string;
  organizationId?: string;
}

function toBase64Url(value: string): string {
  return Buffer.from(value, "utf8").toString("base64url");
}

function fromBase64Url(value: string): string {
  return Buffer.from(value, "base64url").toString("utf8");
}

export function signState(payload: OAuthStatePayload, secret: string): string {
  const encodedPayload = toBase64Url(JSON.stringify(payload));
  const signature = crypto.createHmac("sha256", secret).update(encodedPayload).digest("base64url");
  return `${encodedPayload}.${signature}`;
}

export function verifyState(token: string, secret: string): OAuthStatePayload {
  const [encodedPayload, providedSig] = token.split(".");
  if (!encodedPayload || !providedSig) {
    throw new Error("Invalid OAuth state format");
  }

  const expectedSig = crypto.createHmac("sha256", secret).update(encodedPayload).digest("base64url");

  const a = Buffer.from(providedSig);
  const b = Buffer.from(expectedSig);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) {
    throw new Error("Invalid OAuth state signature");
  }

  const parsed = JSON.parse(fromBase64Url(encodedPayload)) as OAuthStatePayload;
  if (!parsed.nonce || !parsed.requestedAt) {
    throw new Error("Invalid OAuth state payload");
  }

  return parsed;
}
