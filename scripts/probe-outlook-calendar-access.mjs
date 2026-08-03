/**
 * Probe Outlook targets: info@ mailbox vs PAVE Operations group calendar.
 * Usage: node --env-file=.env.local scripts/probe-outlook-calendar-access.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing ${name}`);
  return value;
}

function getClient() {
  const credential = new ClientSecretCredential(
    requireEnv("AZURE_TENANT_ID"),
    requireEnv("AZURE_CLIENT_ID"),
    requireEnv("AZURE_CLIENT_SECRET"),
  );
  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });
  return Client.initWithMiddleware({ authProvider });
}

async function tryGet(label, fn) {
  try {
    const data = await fn();
    console.log(`✅ ${label}`);
    return { ok: true, data };
  } catch (error) {
    const body = error?.body ?? error?.message ?? String(error);
    console.log(`❌ ${label}`);
    console.log(`   ${typeof body === "string" ? body.slice(0, 300) : JSON.stringify(body).slice(0, 300)}`);
    return { ok: false, error: body };
  }
}

async function main() {
  const client = getClient();
  const userId = process.env.OUTLOOK_USER_ID?.trim() || "info@pavetraining.co.uk";
  const groupId =
    process.env.OUTLOOK_GROUP_ID?.trim() ||
    "57c6d553-8503-4f18-b7b6-206904a00976";

  console.log("App:", process.env.AZURE_CLIENT_ID);
  console.log("User target:", userId);
  console.log("Group target:", groupId);
  console.log("");

  // Decode token roles if possible
  const credential = new ClientSecretCredential(
    requireEnv("AZURE_TENANT_ID"),
    requireEnv("AZURE_CLIENT_ID"),
    requireEnv("AZURE_CLIENT_SECRET"),
  );
  const token = await credential.getToken("https://graph.microsoft.com/.default");
  const payload = JSON.parse(
    Buffer.from(token.token.split(".")[1], "base64url").toString("utf8"),
  );
  const roles = payload.roles ?? [];
  console.log("App roles (permissions) on token:");
  for (const role of roles.sort()) console.log(" -", role);
  console.log("");

  await tryGet(`GET /users/${userId} (mailbox exists)`, () =>
    client.api(`/users/${encodeURIComponent(userId)}`).select("id,displayName,mail").get(),
  );

  await tryGet(`GET /users/${userId}/events (app can write/read user calendar)`, () =>
    client
      .api(`/users/${encodeURIComponent(userId)}/events`)
      .top(1)
      .select("id,subject")
      .get(),
  );

  await tryGet(`GET /groups/${groupId} (group exists)`, () =>
    client
      .api(`/groups/${groupId}`)
      .select("id,displayName,mail,groupTypes")
      .get(),
  );

  await tryGet(`GET /groups/${groupId}/events (group calendar — needs app support)`, () =>
    client
      .api(`/groups/${groupId}/events`)
      .top(1)
      .select("id,subject")
      .get(),
  );

  await tryGet(`GET /groups/${groupId}/calendar (group calendar object)`, () =>
    client.api(`/groups/${groupId}/calendar`).get(),
  );

  console.log("\n--- Summary ---");
  console.log(
    "Microsoft Graph: creating/reading M365 *group* calendar events with Application permissions is NOT supported.",
  );
  console.log(
    "Working pattern for this portal (app-only): sync to a user/shared mailbox calendar (e.g. info@) with Calendars.ReadWrite.",
  );
  console.log(
    "Group calendar only works with Delegated auth (signed-in user) + Group.ReadWrite.All.",
  );
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
