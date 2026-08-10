import "server-only";

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";

let cachedClient: Client | null = null;

const GRAPH_READ_MAX_ATTEMPTS = 3;
const GRAPH_MUTATION_MAX_ATTEMPTS = 4;
const GRAPH_RETRY_BASE_MS = 300;

function errorStatus(error: unknown): number | null {
  if (!error || typeof error !== "object") return null;
  const value = (error as { statusCode?: unknown; status?: unknown }).statusCode ??
    (error as { status?: unknown }).status;
  return typeof value === "number" ? value : null;
}

function errorText(error: unknown): string {
  if (error instanceof Error) {
    const cause = "cause" in error ? String(error.cause ?? "") : "";
    return `${error.message} ${cause}`.toLowerCase();
  }
  return String(error ?? "").toLowerCase();
}

/** True for transient network / throttling failures Graph SDK surfaces poorly. */
export function isTransientGraphError(error: unknown): boolean {
  const status = errorStatus(error);
  if (status === -1 || status === 408 || status === 425 || status === 429) {
    return true;
  }
  if (status !== null && status >= 500) return true;

  const text = errorText(error);
  return [
    "fetch failed",
    "terminated",
    "econnreset",
    "econnrefused",
    "etimedout",
    "socket hang up",
    "und_err_",
    "network",
    "other side closed",
  ].some((marker) => text.includes(marker));
}

/** @deprecated Prefer isTransientGraphError — kept for existing import sites. */
export function isTransientGraphReadError(error: unknown): boolean {
  return isTransientGraphError(error);
}

function retryDelay(attempt: number, baseMs = GRAPH_RETRY_BASE_MS): number {
  const jitter = Math.floor(Math.random() * 120);
  return baseMs * 2 ** (attempt - 1) + jitter;
}

async function withGraphRetry<T>(
  operation: () => Promise<T>,
  maxAttempts: number,
  label: string,
): Promise<T> {
  let lastError: unknown;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      if (attempt === maxAttempts || !isTransientGraphError(error)) {
        throw error;
      }
      const waitMs = retryDelay(attempt);
      console.warn(
        `[graph] transient ${label} failure attempt=${attempt}/${maxAttempts} retryInMs=${waitMs} message=${
          error instanceof Error ? error.message : String(error)
        }`,
      );
      await new Promise<void>((resolve) => {
        setTimeout(resolve, waitMs);
      });
    }
  }
  throw lastError;
}

/**
 * Adds a small retry window around Graph GETs for transient network/stream
 * failures that the SDK's HTTP-status retry middleware cannot always see.
 */
export async function withGraphReadRetry<T>(
  operation: () => Promise<T>,
): Promise<T> {
  return withGraphRetry(operation, GRAPH_READ_MAX_ATTEMPTS, "read");
}

/**
 * Retries Graph POST/PATCH/DELETE for the same transient class as reads.
 * Callers must only wrap idempotent-enough operations (or accept rare
 * duplicate creates under extreme failure — matrix upsert prefers update).
 */
export async function withGraphMutationRetry<T>(
  operation: () => Promise<T>,
): Promise<T> {
  return withGraphRetry(operation, GRAPH_MUTATION_MAX_ATTEMPTS, "mutation");
}

/** Richer message for bulk logs when Graph drops the socket. */
export function formatGraphError(error: unknown): string {
  if (!(error instanceof Error)) return String(error ?? "Unknown Graph error");
  const status = errorStatus(error);
  const cause =
    "cause" in error && error.cause
      ? error.cause instanceof Error
        ? error.cause.message
        : String(error.cause)
      : "";
  const parts = [error.message];
  if (status != null) parts.push(`status=${status}`);
  if (cause) parts.push(`cause=${cause}`);
  if (isTransientGraphError(error)) {
    parts.push(
      "(transient network to Microsoft Graph — retry or check VPN/firewall)",
    );
  }
  return parts.join(" ");
}

function getAzureCredentials() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      "Missing Azure Graph credentials. Set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET.",
    );
  }

  return new ClientSecretCredential(tenantId, clientId, clientSecret);
}

/**
 * Server-only Microsoft Graph client using app-only (client credentials) auth.
 * Never import this module from client components.
 */
export function getGraphClient(): Client {
  if (cachedClient) {
    return cachedClient;
  }

  const credential = getAzureCredentials();
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });

  cachedClient = Client.initWithMiddleware({ authProvider });
  return cachedClient;
}
