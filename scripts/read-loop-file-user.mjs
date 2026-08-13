#!/usr/bin/env node
/**
 * Delegated (user) sign-in variant of read-loop-file.mjs.
 * Uses device code flow with the Azure CLI's well-known public client id, so
 * no changes are needed on the tenant's app registration and no admin consent
 * is required (Azure CLI is pre-consented in most tenants).
 *
 * The token has the *user's* rights, not the service-principal's — you'll see
 * exactly what your account sees on the Loop / SharePoint URL.
 *
 * Usage:
 *   node scripts/read-loop-file-user.mjs "<share url>"
 *   node scripts/read-loop-file-user.mjs "<share url>" --save out.bin
 */
import { readFileSync, writeFileSync } from "node:fs";
import { DeviceCodeCredential } from "@azure/identity";

const AZURE_CLI_CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46";

const args = process.argv.slice(2);
const shareUrl = args.find((a) => !a.startsWith("--"));
const saveIdx = args.indexOf("--save");
const savePath = saveIdx >= 0 ? args[saveIdx + 1] : null;

if (!shareUrl) {
  console.error("Usage: node scripts/read-loop-file-user.mjs <shareUrl> [--save file]");
  process.exit(2);
}

function loadDotenv(path) {
  try {
    const text = readFileSync(path, "utf8");
    for (const rawLine of text.split(/\r?\n/)) {
      const line = rawLine.replace(/^\s+/, "");
      if (!line || line.startsWith("#")) continue;
      const eq = line.indexOf("=");
      if (eq <= 0) continue;
      const key = line.slice(0, eq).trim();
      let val = line.slice(eq + 1);
      if (
        (val.startsWith('"') && val.endsWith('"')) ||
        (val.startsWith("'") && val.endsWith("'"))
      ) {
        val = val.slice(1, -1);
      }
      if (!process.env[key]) process.env[key] = val;
    }
  } catch {
    /* env optional */
  }
}
loadDotenv(".env.local");

const tenantId = process.env.AZURE_TENANT_ID || "common";

const credential = new DeviceCodeCredential({
  tenantId,
  clientId: AZURE_CLI_CLIENT_ID,
  userPromptCallback: (info) => {
    console.log("\n============================================================");
    console.log("SIGN-IN REQUIRED");
    console.log("------------------------------------------------------------");
    console.log(`1. Open:  ${info.verificationUri}`);
    console.log(`2. Enter code:  ${info.userCode}`);
    console.log("3. Sign in with the Microsoft account that can view the Loop page.");
    console.log("============================================================\n");
  },
});

async function getToken() {
  const token = await credential.getToken([
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
  ]);
  if (!token?.token) throw new Error("Failed to get Graph token");
  return token.token;
}

function encodeShareUrl(url) {
  const b64 = Buffer.from(url, "utf8").toString("base64");
  const b64url = b64.replace(/=+$/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  return `u!${b64url}`;
}

async function graphFetch(url, token, opts = {}) {
  return fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...(opts.headers || {}),
    },
    redirect: "follow",
  });
}

function preview(buf, max = 1200) {
  const slice = buf.subarray(0, max);
  const hex = slice.toString("hex").match(/.{1,32}/g)?.join("\n") ?? "";
  const utf8 = slice.toString("utf8").replace(/[\x00-\x08\x0E-\x1F]/g, ".");
  return { hex, utf8 };
}

async function main() {
  console.log("Requesting user token (device code flow)…");
  const token = await getToken();
  console.log("Token acquired.\n");

  const shareId = encodeShareUrl(shareUrl);
  const shareEndpoint = `https://graph.microsoft.com/v1.0/shares/${shareId}`;

  console.log("Resolving share via /shares…");
  const shareRes = await graphFetch(shareEndpoint, token);
  if (!shareRes.ok) {
    console.error(`shares → ${shareRes.status} ${shareRes.statusText}`);
    console.error(await shareRes.text());
    process.exit(1);
  }
  const shareJson = await shareRes.json();
  console.log("Share metadata:");
  console.log(JSON.stringify(shareJson, null, 2));
  console.log();

  console.log("Fetching driveItem…");
  const itemRes = await graphFetch(`${shareEndpoint}/driveItem`, token);
  if (!itemRes.ok) {
    console.error(`shares/driveItem → ${itemRes.status} ${itemRes.statusText}`);
    console.error(await itemRes.text());
    process.exit(1);
  }
  const item = await itemRes.json();
  console.log("Drive item:");
  console.log(
    JSON.stringify(
      {
        id: item.id,
        name: item.name,
        size: item.size,
        file: item.file,
        folder: item.folder,
        webUrl: item.webUrl,
        createdDateTime: item.createdDateTime,
        lastModifiedDateTime: item.lastModifiedDateTime,
        parentReference: item.parentReference,
        "@microsoft.graph.downloadUrl": item["@microsoft.graph.downloadUrl"]
          ? "(present)"
          : undefined,
      },
      null,
      2,
    ),
  );
  console.log();

  const downloadUrl = item["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) {
    console.log("No @microsoft.graph.downloadUrl on this item.");
    return;
  }

  console.log("Downloading content…");
  const contentRes = await fetch(downloadUrl, { redirect: "follow" });
  if (!contentRes.ok) {
    console.error(`download → ${contentRes.status} ${contentRes.statusText}`);
    console.error(await contentRes.text());
    process.exit(1);
  }
  const buf = Buffer.from(await contentRes.arrayBuffer());
  console.log(
    `Downloaded ${buf.length} bytes (content-type: ${contentRes.headers.get("content-type")}).`,
  );

  if (savePath) {
    writeFileSync(savePath, buf);
    console.log(`Saved to ${savePath}.`);
  }

  const { hex, utf8 } = preview(buf);
  console.log("\n---- first 1200 bytes (hex) ----");
  console.log(hex);
  console.log("\n---- first 1200 bytes (utf-8, control chars → .) ----");
  console.log(utf8);
}

main().catch((err) => {
  console.error("Failed:", err?.stack || err);
  process.exit(1);
});
