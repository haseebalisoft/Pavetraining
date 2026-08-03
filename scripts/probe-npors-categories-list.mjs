import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
for (const line of readFileSync(resolve(ROOT, ".env.local"), "utf8").split(
  /\r?\n/,
)) {
  const t = line.trim();
  if (!t || t.startsWith("#")) continue;
  const i = t.indexOf("=");
  if (i < 0) continue;
  const k = t.slice(0, i).trim();
  let v = t.slice(i + 1).trim();
  if (
    (v.startsWith('"') && v.endsWith('"')) ||
    (v.startsWith("'") && v.endsWith("'"))
  ) {
    v = v.slice(1, -1);
  }
  if (!process.env[k]) process.env[k] = v;
}

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing ${name}`);
  return value;
}

function sitePath() {
  const siteId = String(requireEnv("SHAREPOINT_SITE_ID")).replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    return `/sites/${siteId.endsWith(":") ? siteId : `${siteId}:`}`;
  }
  return `/sites/${siteId}`;
}

const body = new URLSearchParams({
  client_id: requireEnv("AZURE_CLIENT_ID"),
  client_secret: requireEnv("AZURE_CLIENT_SECRET"),
  scope: "https://graph.microsoft.com/.default",
  grant_type: "client_credentials",
});
const tokRes = await fetch(
  `https://login.microsoftonline.com/${requireEnv("AZURE_TENANT_ID")}/oauth2/v2.0/token`,
  {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  },
);
const token = (await tokRes.json()).access_token;
const root = sitePath();
const listId = requireEnv("SHAREPOINT_NPORS_CATEGORIES_LIST_ID");

const items = [];
let next = `https://graph.microsoft.com/v1.0${root}/lists/${listId}/items?$expand=fields&$top=200`;
while (next) {
  const res = await fetch(next, {
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
    },
  });
  const json = await res.json();
  if (!res.ok) {
    console.error(JSON.stringify(json, null, 2));
    process.exit(1);
  }
  items.push(...(json.value || []));
  next = json["@odata.nextLink"] || null;
}

console.log(`items=${items.length}`);
for (const item of items.slice(0, 20)) {
  const f = item.fields || {};
  console.log(
    [
      f.Title,
      f.Category_x0020_Code,
      f.Machine_x0020_Type,
      f.Category_x0020_Group,
      f.Active,
      f.Customer_x0020_Visible,
      f.Display_x0020_Order,
    ].join(" | "),
  );
}
