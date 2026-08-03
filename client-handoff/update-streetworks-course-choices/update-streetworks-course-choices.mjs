/**
 * PAVE Training — set NRSWA Register Course choices (Site Owner).
 *
 * Run from this folder:
 *   npm install
 *   node update-streetworks-course-choices.mjs --dry-run
 *   node update-streetworks-course-choices.mjs
 */
import { readFileSync } from "node:fs";
import { resolve, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { DeviceCodeCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const HERE = dirname(fileURLToPath(import.meta.url));
const DRY_RUN = process.argv.includes("--dry-run");

const PUBLIC_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const TENANT_ID = "a74de0ea-7d76-44f4-9cdb-9ecf22b84079";
const SITE_PATH =
  "/sites/pavetraining.sharepoint.com:/sites/PaveTrainingOperationAdmin:";
const NRSWA_LIST_ID = "9bdcf9b0-31f9-45b9-b3c8-29a699482eb4";

const TARGET = [
  "Operative",
  "Supervisor",
  "Operative Reassessment",
  "Supervisor Reassessment",
];

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
  // Keep txt in pack for humans; script uses TARGET constants.
  void readFileSync(
    resolve(HERE, "Streetworks-Course-SharePoint-Choices.txt"),
    "utf8",
  );

  console.log("PAVE — Update Streetworks Course choices");
  console.log(`Target: ${TARGET.join(" · ")}`);
  console.log(DRY_RUN ? "\nDRY RUN\n" : "\nLIVE RUN\n");
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
    `${SITE_PATH}/lists/${NRSWA_LIST_ID}/columns?$top=200`,
  );
  const column = columns.find(
    (c) =>
      String(c.name || "").toLowerCase() === "course" ||
      String(c.displayName || "").toLowerCase() === "course",
  );
  if (!column) {
    throw new Error('Could not find "Course" column on NRSWA Register.');
  }

  const current = column.choice?.choices || [];
  console.log(`Column: ${column.displayName} (${column.name})`);
  console.log(`Current (${current.length}): ${current.join(", ")}`);

  if (DRY_RUN) {
    console.log("\nDry run OK. Run again without --dry-run to apply.");
    return;
  }

  await client
    .api(`${SITE_PATH}/lists/${NRSWA_LIST_ID}/columns/${column.id}`)
    .patch({
      choice: {
        ...(column.choice || {}),
        allowTextEntry: false,
        choices: TARGET,
      },
    });

  console.log("\nDone. Course choices updated to the 4 Streetworks options.");
}

main().catch((error) => {
  console.error("\nFAILED:", error?.message || error);
  process.exit(1);
});
