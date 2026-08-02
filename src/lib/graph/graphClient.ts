import "server-only";

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";

let cachedClient: Client | null = null;

const GRAPH_READ_MAX_ATTEMPTS = 3;
const GRAPH_READ_RETRY_BASE_MS = 300;

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

/** True only for failures where repeating an idempotent Graph GET is safe. */
export function isTransientGraphReadError(error: unknown): boolean {
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
  ].some((marker) => text.includes(marker));
}

function retryDelay(attempt: number): number {
  return GRAPH_READ_RETRY_BASE_MS * 2 ** (attempt - 1);
}

/**
 * Adds a small retry window around Graph GETs for transient network/stream
 * failures that the SDK's HTTP-status retry middleware cannot always see.
 */
export async function withGraphReadRetry<T>(
  operation: () => Promise<T>,
): Promise<T> {
  let lastError: unknown;
  for (let attempt = 1; attempt <= GRAPH_READ_MAX_ATTEMPTS; attempt += 1) {
    try {
      return await operation();
    } catch (error) {
      lastError = error;
      if (
        attempt === GRAPH_READ_MAX_ATTEMPTS ||
        !isTransientGraphReadError(error)
      ) {
        throw error;
      }
      await new Promise<void>((resolve) => {
        setTimeout(resolve, retryDelay(attempt));
      });
    }
  }
  throw lastError;
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
