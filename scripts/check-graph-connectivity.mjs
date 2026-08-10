/**
 * Quick Graph connectivity check using .env.local (no secrets printed).
 *   node scripts/check-graph-connectivity.mjs
 */
import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { ClientSecretCredential } = require("@azure/identity");

function loadEnvLocal() {
  const path = resolve(process.cwd(), ".env.local");
  if (!existsSync(path)) throw new Error(".env.local not found");
  const env = {};
  for (const line of readFileSync(path, "utf8").split(/\r?\n/)) {
    if (!line || line.startsWith("#")) continue;
    const i = line.indexOf("=");
    if (i < 0) continue;
    let v = line.slice(i + 1).trim();
    if (
      (v.startsWith('"') && v.endsWith('"')) ||
      (v.startsWith("'") && v.endsWith("'"))
    ) {
      v = v.slice(1, -1);
    }
    env[line.slice(0, i).trim()] = v;
  }
  return env;
}

const env = loadEnvLocal();
const needed = [
  "AZURE_TENANT_ID",
  "AZURE_CLIENT_ID",
  "AZURE_CLIENT_SECRET",
  "SHAREPOINT_SITE_ID",
  "SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID",
];
for (const k of needed) {
  console.log(k, env[k] ? "SET" : "MISSING");
}

const credential = new ClientSecretCredential(
  env.AZURE_TENANT_ID,
  env.AZURE_CLIENT_ID,
  env.AZURE_CLIENT_SECRET,
);

const t0 = Date.now();
try {
  const token = await credential.getToken(
    "https://graph.microsoft.com/.default",
  );
  console.log("token ok ms=", Date.now() - t0);

  // Match getSharePointSiteApiRoot(): path IDs (host:/sites/Name) need a trailing colon.
  let siteId = String(env.SHAREPOINT_SITE_ID).replace(/\/+$/, "");
  if (siteId.includes(":/") && !siteId.endsWith(":")) siteId = `${siteId}:`;
  const siteRoot = `/sites/${siteId}`;
  console.log("site path style", siteId.includes(":/") ? "host:/sites/..." : "composite/id");

  const siteRes = await fetch(`https://graph.microsoft.com/v1.0${siteRoot}`, {
    headers: { Authorization: `Bearer ${token.token}` },
  });
  console.log("site status", siteRes.status, "ms", Date.now() - t0);

  const listId = env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID;
  const listUrl = `https://graph.microsoft.com/v1.0${siteRoot}/lists/${listId}?$select=id,displayName`;
  const listRes = await fetch(listUrl, {
    headers: { Authorization: `Bearer ${token.token}` },
  });
  console.log("matrix example list status", listRes.status, "ms", Date.now() - t0);
  const listJson = await listRes.json();
  console.log("list displayName", listJson.displayName ?? listJson.error?.message);

  const itemsUrl = `https://graph.microsoft.com/v1.0${siteRoot}/lists/${listId}/items?$top=1&$select=id`;
  const itemsRes = await fetch(itemsUrl, {
    headers: { Authorization: `Bearer ${token.token}` },
  });
  console.log("items $top=1 status", itemsRes.status, "ms", Date.now() - t0);

  // Lightweight create+delete smoke (proves writes work). Skip if SMOKE_WRITE!=1.
  if (process.env.SMOKE_WRITE === "1") {
    const createUrl = `https://graph.microsoft.com/v1.0${siteRoot}/lists/${listId}/items`;
    const createRes = await fetch(createUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token.token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        fields: { Title: `__connectivity_probe_${Date.now()}` },
      }),
    });
    console.log("create smoke status", createRes.status, "ms", Date.now() - t0);
    const created = await createRes.json();
    if (created.id) {
      const del = await fetch(
        `https://graph.microsoft.com/v1.0${siteRoot}/lists/${listId}/items/${created.id}`,
        {
          method: "DELETE",
          headers: { Authorization: `Bearer ${token.token}` },
        },
      );
      console.log("delete smoke status", del.status, "ms", Date.now() - t0);
    } else {
      console.log("create error", created.error?.message ?? created);
    }
  }
} catch (error) {
  console.error("FAIL", error?.message ?? error);
  if (error?.cause) console.error("cause", error.cause);
  process.exitCode = 1;
}
