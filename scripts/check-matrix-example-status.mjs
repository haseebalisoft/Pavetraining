import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const XLSX = require("xlsx");
const { ClientSecretCredential } = require("@azure/identity");

function loadEnv() {
  const env = {};
  for (const line of readFileSync(".env.local", "utf8").split(/\r?\n/)) {
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

const nk = (v) =>
  String(v ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");

const env = loadEnv();
const wb = XLSX.read(readFileSync("Training matrix example.xlsx"), {
  type: "buffer",
  raw: true,
});
const m = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {
  header: 1,
  defval: null,
  raw: true,
  blankrows: false,
});
const headers = m[0].map((c) => (c == null ? "" : String(c).trim()));
const ni = headers.findIndex((h) => /^name$/i.test(h));
const targets = new Set();
for (let i = 1; i < m.length; i += 1) {
  const k = nk(m[i][ni]);
  if (k) targets.add(k);
}

let site = env.SHAREPOINT_SITE_ID.replace(/\/+$/, "");
if (site.includes(":/") && !site.endsWith(":")) site += ":";
const cred = new ClientSecretCredential(
  env.AZURE_TENANT_ID,
  env.AZURE_CLIENT_ID,
  env.AZURE_CLIENT_SECRET,
);
const t = await cred.getToken("https://graph.microsoft.com/.default");
const auth = { Authorization: `Bearer ${t.token}` };
const root = `https://graph.microsoft.com/v1.0/sites/${site}`;
const listId = env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID;
let url = `${root}/lists/${listId}/items?$expand=fields&$top=200`;
const items = [];
while (url) {
  const r = await fetch(url, { headers: auth });
  const j = await r.json();
  items.push(...(j.value || []));
  url = j["@odata.nextLink"] || "";
}

const found = [];
const missing = new Set(targets);
for (const it of items) {
  const title = nk(it.fields?.Title);
  if (!targets.has(title)) continue;
  found.push({
    title: it.fields.Title,
    id: it.id,
    status: it.fields.MatrixLinkStatus || "(blank)",
    wf: it.fields.WorkforceItemId ?? null,
  });
  missing.delete(title);
}
const counts = {};
for (const r of found) counts[r.status] = (counts[r.status] || 0) + 1;
console.log({
  targets: targets.size,
  found: found.length,
  missing: [...missing],
  statuses: counts,
});
