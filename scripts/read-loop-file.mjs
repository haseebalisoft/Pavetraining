#!/usr/bin/env node
/**
 * Resolve a SharePoint / Loop share URL to a driveItem via Microsoft Graph
 * (app-only auth) and dump metadata + raw file content.
 *
 * Loop files are Fluid Framework blobs; the raw content is not human-readable
 * but we can still tell you the driveItem type, size, mime, and dump a hex/UTF-8
 * preview so you can decide next steps.
 *
 * Usage:
 *   node scripts/read-loop-file.mjs "<share url>"
 *   node scripts/read-loop-file.mjs "<share url>" --save out.bin
 */
import { readFileSync, writeFileSync } from "node:fs";
import { ClientSecretCredential } from "@azure/identity";

const args = process.argv.slice(2);
const shareUrl = args.find((a) => !a.startsWith("--"));
const saveIdx = args.indexOf("--save");
const savePath = saveIdx >= 0 ? args[saveIdx + 1] : null;

if (!shareUrl) {
  console.error("Usage: node scripts/read-loop-file.mjs <shareUrl> [--save file]");
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
    /* file missing – rely on real env */
  }
}
loadDotenv(".env.local");

const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;
if (!tenantId || !clientId || !clientSecret) {
  console.error("Missing AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET.");
  process.exit(2);
}

const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

async function getToken() {
  const token = await credential.getToken("https://graph.microsoft.com/.default");
  if (!token?.token) throw new Error("Failed to get Graph token");
  return token.token;
}

/** Encode share URL per docs: u! + base64url without padding. */
function encodeShareUrl(url) {
  const b64 = Buffer.from(url, "utf8").toString("base64");
  const b64url = b64.replace(/=+$/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  return `u!${b64url}`;
}

async function graphFetch(url, token, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      ...(opts.headers || {}),
    },
    redirect: "follow",
  });
  return res;
}

function preview(buf, max = 800) {
  const slice = buf.subarray(0, max);
  const hex = slice.toString("hex").match(/.{1,32}/g)?.join("\n") ?? "";
  const utf8 = slice.toString("utf8").replace(/[\x00-\x08\x0E-\x1F]/g, ".");
  return { hex, utf8 };
}

async function main() {
  console.log("Authenticating…");
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
    console.log(
      "No @microsoft.graph.downloadUrl on this item (likely a Loop container / folder).",
    );
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
  console.log(`Downloaded ${buf.length} bytes (content-type: ${contentRes.headers.get("content-type")}).`);

  if (savePath) {
    writeFileSync(savePath, buf);
    console.log(`Saved to ${savePath}.`);
  }

  const { hex, utf8 } = preview(buf);
  console.log("\n---- first 800 bytes (hex) ----");
  console.log(hex);
  console.log("\n---- first 800 bytes (utf-8, control chars → .) ----");
  console.log(utf8);
}

main().catch((err) => {
  console.error("Failed:", err?.stack || err);
  process.exit(1);
});
