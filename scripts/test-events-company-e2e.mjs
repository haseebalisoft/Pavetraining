/**
 * End-to-end Graph test for Events company lookup + DoNotSync payload.
 *
 * Usage:
 *   node --env-file=.env.local scripts/test-events-company-e2e.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
  return value;
}

function siteRoot() {
  const siteId = String(requireEnv("SHAREPOINT_SITE_ID")).replace(/\/+$/, "");
  if (siteId.includes(":/")) {
    return `/sites/${siteId.endsWith(":") ? siteId : `${siteId}:`}`;
  }
  return `/sites/${siteId}`;
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

function stripSharePointHtml(value) {
  if (value == null) return null;
  const raw = String(value).trim();
  if (!raw) return null;
  if (!raw.includes("<") && !raw.includes("&")) return raw;
  return raw
    .replace(/<\s*br\s*\/?\s*>/gi, "\n")
    .replace(/<\/\s*p\s*>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&#160;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}

const results = [];
function pass(name, detail = "") {
  results.push({ ok: true, name, detail });
  console.log(`PASS  ${name}${detail ? ` — ${detail}` : ""}`);
}
function fail(name, detail = "") {
  results.push({ ok: false, name, detail });
  console.log(`FAIL  ${name}${detail ? ` — ${detail}` : ""}`);
}

async function main() {
  const client = getClient();
  const eventsList = requireEnv("SHAREPOINT_EVENTS_LIST_ID");
  const companyList = requireEnv("SHAREPOINT_COMPANY_LIST_ID");
  const root = `${siteRoot()}/lists/${eventsList}/items`;
  const colsRoot = `${siteRoot()}/lists/${eventsList}/columns`;

  const cols = await client.api(colsRoot).top(300).get();
  const doNotSyncCol = (cols.value || []).find((c) => c.name === "DoNotSync");
  const eventCompanyCol = (cols.value || []).find((c) => c.name === "EventCompany");

  if (doNotSyncCol?.text && !doNotSyncCol?.boolean) {
    pass("DoNotSync is text column");
  } else {
    fail("DoNotSync column type unexpected", JSON.stringify(doNotSyncCol));
  }

  if (eventCompanyCol?.lookup?.listId === companyList) {
    pass("EventCompany lookup points at Company List");
  } else {
    fail(
      "EventCompany lookup target",
      String(eventCompanyCol?.lookup?.listId),
    );
  }

  const uniqueOn = Boolean(eventCompanyCol?.enforceUniqueValues);
  console.log(`INFO  EventCompany.enforceUniqueValues = ${uniqueOn}`);

  // Cleanup leftover probe items from earlier runs.
  const existing = await client
    .api(root)
    .expand("fields")
    .top(200)
    .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
    .get();

  for (const item of existing.value || []) {
    const title = String(item.fields?.Title || "");
    if (
      /^(E2E |NEG-|Probe)/i.test(title) ||
      /ProbeFix|ProbeFixed|Cal UID|CalCo|Cal /i.test(title)
    ) {
      await client.api(`${root}/${item.id}`).delete().catch(() => {});
      console.log(`INFO  deleted leftover ${item.id} ${title}`);
    }
  }

  // Ensure seed bookings are linked.
  await client
    .api(`${root}/7/fields`)
    .patch({
      EventCompanyLookupId: 27,
      DoNotSync: "Yes",
      Customer_x0020_Visible: true,
    })
    .catch((e) => console.log("INFO  patch item 7:", e.message));
  await client
    .api(`${root}/8/fields`)
    .patch({
      EventCompanyLookupId: 34,
      DoNotSync: "Yes",
      Customer_x0020_Visible: true,
    })
    .catch((e) => console.log("INFO  patch item 8:", e.message));

  const stamp = Date.now();
  const start = new Date(Date.now() + 95 * 86400000).toISOString();
  const end = new Date(Date.now() + 95 * 86400000 + 2 * 3600000).toISOString();

  // A) Boolean DoNotSync must fail (regression for the 500 bug).
  try {
    await client.api(root).post({
      fields: {
        Title: `NEG-bool-${stamp}`,
        EventDate: start,
        EndDate: end,
        EventCompanyLookupId: 27,
        Customer_x0020_Visible: true,
        DoNotSync: false,
        SyncStatus: "Pending",
        SyncDirection: "SharePointToOutlook",
        LastSyncSource: "SharePoint",
      },
    });
    fail("Boolean DoNotSync should be rejected");
  } catch (error) {
    if (
      /generalException/i.test(error.code || "") ||
      /general exception/i.test(error.message || "")
    ) {
      pass("Boolean DoNotSync correctly rejected");
    } else {
      pass("Boolean DoNotSync rejected", `${error.code} ${error.message}`);
    }
  }

  // Free Murphy unique slot, create with fixed app-like payload, verify, delete, restore.
  await client.api(`${root}/7/fields`).patch({ EventCompanyLookupId: null });

  let createdId = null;
  try {
    const created = await client.api(root).post({
      fields: {
        Title: `E2E Murphy Create ${stamp}`,
        EventDate: start,
        EndDate: end,
        EventCompanyLookupId: 27,
        Customer_x0020_Visible: true,
        TrainingAddress: "Full training address",
        Location: "Murphy Plant Site",
        Description: "E2E test event",
        DoNotSync: "No",
        SyncStatus: "Pending",
        SyncDirection: "SharePointToOutlook",
        LastSyncSource: "SharePoint",
      },
    });
    createdId = String(created.id);
    const fields = created.fields || {};

    if (String(fields.EventCompanyLookupId) === "27") {
      pass("Create saves EventCompanyLookupId=27", `id=${createdId}`);
    } else {
      fail("Create EventCompanyLookupId", String(fields.EventCompanyLookupId));
    }

    if (fields.DoNotSync === "No") {
      pass("DoNotSync stored as text No");
    } else {
      fail("DoNotSync value", String(fields.DoNotSync));
    }

    if (fields.Customer_x0020_Visible === true) {
      pass("Customer Visible = true");
    } else {
      fail("Customer Visible", String(fields.Customer_x0020_Visible));
    }
  } catch (error) {
    fail("Create fixed payload", `${error.statusCode} ${error.message}`);
  }

  if (createdId) {
    const item = await client.api(`${root}/${createdId}`).expand("fields").get();
    const companies = await client
      .api(`${siteRoot()}/lists/${companyList}/items`)
      .expand("fields")
      .top(200)
      .get();
    const company = (companies.value || []).find(
      (row) => String(row.id) === String(item.fields?.EventCompanyLookupId),
    );
    const companyName = company?.fields?.CompanyName || company?.fields?.Title;
    if (/murphy/i.test(String(companyName || ""))) {
      pass("Admin company name resolve", companyName);
    } else {
      fail("Admin company name resolve", String(companyName));
    }

    const html =
      '<div class="ExternalClassABC"><p>&#160;Full training address<br></p></div>';
    const stripped = stripSharePointHtml(html);
    if (stripped === "Full training address") {
      pass("TrainingAddress HTML strip", stripped);
    } else {
      fail("TrainingAddress HTML strip", String(stripped));
    }

    const allItems = await client
      .api(root)
      .expand("fields")
      .top(200)
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .get();

    const murphyVisible = (allItems.value || []).filter(
      (row) =>
        String(row.fields?.EventCompanyLookupId) === "27" &&
        row.fields?.Customer_x0020_Visible === true,
    );
    const fastVisible = (allItems.value || []).filter(
      (row) =>
        String(row.fields?.EventCompanyLookupId) === "34" &&
        row.fields?.Customer_x0020_Visible === true,
    );

    if (murphyVisible.some((row) => String(row.id) === createdId)) {
      pass(
        "Murphy customer filter includes created event",
        `count=${murphyVisible.length}`,
      );
    } else {
      fail("Murphy customer filter missing event", `count=${murphyVisible.length}`);
    }

    if (!fastVisible.some((row) => String(row.id) === createdId)) {
      pass("Fast customer does not see Murphy event", `count=${fastVisible.length}`);
    } else {
      fail("Fast customer incorrectly sees Murphy event");
    }

    // Delete created item BEFORE restoring seed (unique column).
    await client.api(`${root}/${createdId}`).delete();
    pass("Cleanup deleted test event", createdId);
    createdId = null;
  }

  try {
    await client.api(`${root}/7/fields`).patch({
      EventCompanyLookupId: 27,
      DoNotSync: "Yes",
      Customer_x0020_Visible: true,
    });
    pass("Restored Murphy Test Booking EventCompany=27");
  } catch (error) {
    fail("Restore Murphy Test Booking", error.message);
  }

  // Unique still blocks a second Murphy event while seed holds the slot.
  if (uniqueOn) {
    try {
      await client.api(root).post({
        fields: {
          Title: `E2E Unique ${stamp}`,
          EventDate: start,
          EndDate: end,
          EventCompanyLookupId: 27,
          Customer_x0020_Visible: true,
          DoNotSync: "No",
          SyncStatus: "Pending",
          SyncDirection: "SharePointToOutlook",
          LastSyncSource: "SharePoint",
        },
      });
      fail("Unique constraint should block second Murphy event");
    } catch (error) {
      const body = String(error.body || "");
      if (
        /unique constraints/i.test(error.message || "") ||
        /unique constraints/i.test(body)
      ) {
        pass(
          "Unique constraint still blocks 2nd Murphy event",
          "Turn off Enforce unique values in SharePoint to allow multiple events per company",
        );
      } else {
        fail("Unexpected unique-path error", `${error.statusCode} ${error.message}`);
      }
    }
  } else {
    pass("EventCompany unique already disabled — multiple events per company allowed");
  }

  const finalItems = await client
    .api(root)
    .expand("fields")
    .top(200)
    .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
    .get();
  const item7 = (finalItems.value || []).find((row) => String(row.id) === "7");
  const item8 = (finalItems.value || []).find((row) => String(row.id) === "8");

  if (String(item7?.fields?.EventCompanyLookupId) === "27") {
    pass("Seed item 7 still linked to Murphy");
  } else {
    fail(
      "Seed item 7 company",
      String(item7?.fields?.EventCompanyLookupId),
    );
  }
  if (String(item8?.fields?.EventCompanyLookupId) === "34") {
    pass("Seed item 8 still linked to Fast");
  } else {
    fail(
      "Seed item 8 company",
      String(item8?.fields?.EventCompanyLookupId),
    );
  }

  console.log("\n===== SUMMARY =====");
  const failed = results.filter((row) => !row.ok);
  console.log(`Passed: ${results.length - failed.length}`);
  console.log(`Failed: ${failed.length}`);
  for (const row of failed) {
    console.log(` - ${row.name}: ${row.detail}`);
  }

  if (failed.length) process.exit(1);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
