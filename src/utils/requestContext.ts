import { Request } from "express";

export function getOrganizationIdFromRequest(req: Request): string {
  const fromQuery = typeof req.query.organizationId === "string" ? req.query.organizationId : "";
  const fromBody = typeof req.body?.organizationId === "string" ? req.body.organizationId : "";
  const fromHeader = typeof req.headers["x-organization-id"] === "string" ? req.headers["x-organization-id"] : "";

  const organizationId = String(fromQuery || fromBody || fromHeader || "").trim();

  if (!organizationId) {
    throw new Error("organizationId is required");
  }

  return organizationId;
}
