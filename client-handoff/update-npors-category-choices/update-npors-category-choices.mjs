/**
 * PAVE Training — expand SharePoint NPORS Category choices (Site Owner).
 *
 * No app secret required. You sign in with your Microsoft account.
 *
 * Requirements:
 *   - Node.js 18+ (https://nodejs.org)
 *   - You are a Site Owner on PaveTrainingOperationAdmin
 *
 * Run (from this folder):
 *   npm install
 *   node update-npors-category-choices.mjs --dry-run
 *   node update-npors-category-choices.mjs
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { DeviceCodeCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const HERE = dirname(fileURLToPath(import.meta.url));
const DRY_RUN = process.argv.includes("--dry-run");

// Public Microsoft Graph PowerShell client — device code, no secret to share.
const PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const TENANT_ID = "a74de0ea-7d76-44f4-9cdb-9ecf22b84079"; // pavetraining
const SITE_PATH =
  "/sites/pavetraining.sharepoint.com:/sites/PaveTrainingOperationAdmin:";
const NPORS_LIST_ID = "a0c8ef27-dac2-4a07-92ea-af140c9ca5d6";

function loadChoicesFromTxt() {
  const text = readFileSync(
    resolve(HERE, "NPORS-Category-SharePoint-Choices.txt"),
    "utf8",
  );
  const begin = text.indexOf("--- BEGIN CHOICES ---");
  const end = text.indexOf("--- END CHOICES ---");
  if (begin < 0 || end < 0) {
    throw new Error("Choices markers not found in NPORS-Category-SharePoint-Choices.txt");
  }
  return text
    .slice(begin + "--- BEGIN CHOICES ---".length, end)
    .split(/\r?\n/)
    .map((line) => line.trim().toUpperCase())
    .filter((line) => /^N\d+[A-Z]?$/.test(line) || /^S\d+$/.test(line));
}

async function listAll(client, firstPath) {
  const items = [];
  let path = firstPath;
  while (path) {
    const page = await client.api(path).get();
    items.push(...(page.value || []));
    path = page["@odata.nextLink"]
      ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
      : null;
  }
  return items;
}

async function main() {
  const target = loadChoicesFromTxt();
  if (target.length < 10) {
    throw new Error(`Too few choices loaded (${target.length}). Check the .txt file.`);
  }

  console.log("PAVE — Update NPORS Category SharePoint choices");
  console.log(`Codes to apply: ${target.length}`);
  console.log(
    DRY_RUN
      ? "\nDRY RUN — will not write changes.\n"
      : "\nLIVE RUN — will update SharePoint.\n",
  );
  console.log("Sign in as a SharePoint Site Owner when prompted.\n");

  const credential = new DeviceCodeCredential({
    tenantId: TENANT_ID,
    clientId: PUBLIC_CLIENT_ID,
    userPromptCallback: (info) => {
      console.log("========================================");
      console.log(info.message);
      console.log("========================================\n");
    },
  });

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/Sites.FullControl.All"],
  });
  const client = Client.initWithMiddleware({ authProvider });

  const columns = await listAll(
    client,
    `${SITE_PATH}/lists/${NPORS_LIST_ID}/columns?$top=200`,
  );
  const column = columns.find(
    (c) =>
      String(c.name || "").toLowerCase() === "nporscategory" ||
      /npors\s*category/i.test(c.displayName || ""),
  );
  if (!column) {
    throw new Error('Could not find column "NPORS Category" on NPORS Register.');
  }

  const current = column.choice?.choices || [];
  console.log(`Column: ${column.displayName} (${column.name})`);
  console.log(`Current (${current.length}): ${current.join(", ")}`);
  console.log(`Target  (${target.length}): ${target.slice(0, 8).join(", ")} …`);

  if (DRY_RUN) {
    console.log("\nDry run OK. Run again without --dry-run to apply.");
    return;
  }

  await client
    .api(`${SITE_PATH}/lists/${NPORS_LIST_ID}/columns/${column.id}`)
    .patch({
      choice: {
        ...(column.choice || {}),
        allowTextEntry: false,
        choices: target,
      },
    });

  console.log(`\nDone. NPORS Category now has ${target.length} choices.`);
}

main().catch((error) => {
  console.error("\nFAILED:", error?.message || error);
  console.error(
    "\nYou must be a Site Owner on:\n" +
      "  https://pavetraining.sharepoint.com/sites/PaveTrainingOperationAdmin\n",
  );
  process.exit(1);
});
