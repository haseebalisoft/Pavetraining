/**
 * Align NRSWA Register "Course" choice column with the app:
 *   Operative | Supervisor | Operative Reassessment | Supervisor Reassessment
 *
 *   node --env-file=.env.local scripts/update-streetworks-course-choices.mjs
 *   node --env-file=.env.local scripts/update-streetworks-course-choices.mjs --dry-run
 *
 * Note: App-only Graph often returns 403 on column schema PATCH. If that happens,
 * update the Course choices in SharePoint list settings for NRSWA Register instead.
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const DRY_RUN = process.argv.includes("--dry-run");

const COURSE_CHOICES = [
  "Operative",
  "Supervisor",
  "Operative Reassessment",
  "Supervisor Reassessment",
];

function loadEnv() {
  try {
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
  } catch {
    // ignore
  }
}

loadEnv();

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
}

function sitePath() {
  const siteId = String(requireEnv("SHAREPOINT_SITE_ID")).replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    return `/sites/${siteId.endsWith(":") ? siteId : `${siteId}:`}`;
  }
  return `/sites/${siteId}`;
}

async function getToken() {
  const body = new URLSearchParams({
    client_id: requireEnv("AZURE_CLIENT_ID"),
    client_secret: requireEnv("AZURE_CLIENT_SECRET"),
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await fetch(
    `https://login.microsoftonline.com/${requireEnv("AZURE_TENANT_ID")}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );
  if (!res.ok) throw new Error(`Token ${res.status}: ${await res.text()}`);
  return (await res.json()).access_token;
}

async function graph(token, path, init = {}) {
  const url = path.startsWith("http")
    ? path
    : `https://graph.microsoft.com/v1.0${path}`;
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(init.headers || {}),
    },
  });
  if (!res.ok) {
    throw new Error(`Graph ${res.status} ${url}: ${await res.text()}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function listAll(token, firstPath) {
  const items = [];
  let path = firstPath;
  while (path) {
    const page = await graph(token, path);
    items.push(...(page.value || []));
    path = page["@odata.nextLink"] || null;
  }
  return items;
}

const token = await getToken();
const root = sitePath();
const listId = requireEnv("SHAREPOINT_NRSWA_REGISTER_LIST_ID");

const columns = await listAll(token, `${root}/lists/${listId}/columns?$top=200`);
const courseCol = columns.find(
  (c) =>
    String(c.name || "").toLowerCase() === "course" ||
    String(c.displayName || "").toLowerCase() === "course",
);

if (!courseCol) {
  throw new Error('Could not find "Course" column on NRSWA Register.');
}

const current = courseCol.choice?.choices || [];
console.log(`NRSWA Register Course column: ${courseCol.name} (${courseCol.id})`);
console.log(`Current choices (${current.length}):`);
for (const c of current) console.log(`  - ${c}`);
console.log(`\nTarget choices (${COURSE_CHOICES.length}):`);
for (const c of COURSE_CHOICES) console.log(`  - ${c}`);

const same =
  current.length === COURSE_CHOICES.length &&
  COURSE_CHOICES.every((c, i) => current[i] === c);

if (same) {
  console.log("\nAlready aligned — nothing to do.");
  process.exit(0);
}

if (DRY_RUN) {
  console.log("\n[dry-run] Would PATCH choice.choices to target list.");
  process.exit(0);
}

await graph(token, `${root}/lists/${listId}/columns/${courseCol.id}`, {
  method: "PATCH",
  body: JSON.stringify({
    choice: {
      ...(courseCol.choice || {}),
      allowTextEntry: false,
      choices: COURSE_CHOICES,
    },
  }),
});

console.log("\nUpdated Course choices on NRSWA Register.");
