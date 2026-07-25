import "server-only";

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";

let cachedClient: Client | null = null;

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
