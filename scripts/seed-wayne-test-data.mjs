/**
 * Seeds Wayne's cross-company access test data into SharePoint via Microsoft Graph.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/seed-wayne-test-data.mjs
 */

import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const LISTS = {
  company: process.env.SHAREPOINT_COMPANY_LIST_ID,
  workforce: process.env.SHAREPOINT_WORKFORCE_LIST_ID,
  trainingMatrix: process.env.SHAREPOINT_TRAINING_MATRIX_LIST_ID,
  npors: process.env.SHAREPOINT_NPORS_REGISTER_LIST_ID,
  eusr: process.env.SHAREPOINT_EUSR_REGISTER_LIST_ID,
  nrswa: process.env.SHAREPOINT_NRSWA_REGISTER_LIST_ID,
  inHouse: process.env.SHAREPOINT_IN_HOUSE_CERTIFICATES_LIST_ID,
  nvq: process.env.SHAREPOINT_NVQ_REGISTER_LIST_ID,
  documents: process.env.SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID,
  events: process.env.SHAREPOINT_EVENTS_LIST_ID,
  permissions: process.env.SHAREPOINT_PERMISSIONS_LIST_ID,
};

const SITE = process.env.SHAREPOINT_SITE_ID;

function requireEnv(name) {
  const value = process.env[name];
  if (!value) throw new Error(`Missing env ${name}`);
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

function siteRoot() {
  const siteId = String(SITE).replace(/\/+$/, "");
  // Path-based IDs need a trailing colon before /lists
  if (siteId.includes(":/")) {
    const withTransition = siteId.endsWith(":") ? siteId : `${siteId}:`;
    return `/sites/${withTransition}`;
  }
  return `/sites/${siteId}`;
}

function escapeOData(value) {
  return String(value).replace(/'/g, "''");
}

async function listItems(client, listId, filter, top = 200) {
  let request = client
    .api(`${siteRoot()}/lists/${listId}/items`)
    .expand("fields")
    .top(top)
    .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
  if (filter) request = request.filter(filter);
  const res = await request.get();
  return res.value ?? [];
}

async function createItem(client, listId, fields) {
  try {
    const created = await client
      .api(`${siteRoot()}/lists/${listId}/items`)
      .post({ fields });
    return { id: String(created.id), fields: created.fields ?? fields };
  } catch (error) {
    const body =
      error?.body ?? error?.message ?? JSON.stringify(error, null, 2);
    const enriched = new Error(
      `createItem failed: ${typeof body === "string" ? body : JSON.stringify(body)}`,
    );
    enriched.cause = error;
    throw enriched;
  }
}

async function updateItemFields(client, listId, itemId, fields) {
  await client
    .api(`${siteRoot()}/lists/${listId}/items/${itemId}/fields`)
    .patch(fields);
}

async function findCompanyByName(client, name) {
  const items = await listItems(
    client,
    LISTS.company,
    `fields/CompanyName eq '${escapeOData(name)}'`,
  );
  if (items.length > 0) {
    return {
      id: String(items[0].id),
      name: items[0].fields?.CompanyName ?? name,
    };
  }
  // Fallback scan (spaces / casing quirks)
  const all = await listItems(client, LISTS.company, undefined, 5000);
  const match = all.find(
    (item) =>
      String(item.fields?.CompanyName ?? "")
        .trim()
        .toLowerCase() === name.trim().toLowerCase(),
  );
  if (!match) return null;
  return { id: String(match.id), name: match.fields?.CompanyName ?? name };
}

async function ensureCompany(client, name, number) {
  const existing = await findCompanyByName(client, name);
  if (existing) {
    console.log(`  company exists: ${existing.name} (#${existing.id})`);
    return existing;
  }
  const created = await createItem(client, LISTS.company, {
    Title: name,
    CompanyName: name,
    CompanyNumber: number,
    Status: "Active",
  });
  console.log(`  company created: ${name} (#${created.id})`);
  return { id: created.id, name };
}

async function findWorkforce(client, candidateName, companyId) {
  try {
    const filtered = await listItems(
      client,
      LISTS.workforce,
      `fields/CandidateName eq '${escapeOData(candidateName)}'`,
      50,
    );
    const match = filtered.find(
      (item) =>
        String(item.fields?.CompanyNameLookupId ?? "").trim() ===
        String(companyId),
    );
    if (match) return match;
  } catch {
    // Fall through
  }
  const all = await listItems(client, LISTS.workforce, undefined, 200);
  return (
    all.find((item) => {
      const n = String(item.fields?.CandidateName ?? "").trim().toLowerCase();
      const c = String(item.fields?.CompanyNameLookupId ?? "").trim();
      return (
        n === candidateName.trim().toLowerCase() && c === String(companyId)
      );
    }) ?? null
  );
}

async function ensureWorkforce(client, row) {
  const existing = await findWorkforce(
    client,
    row.candidateName,
    row.companyId,
  );
  if (existing) {
    console.log(
      `  workforce exists: ${row.candidateName} (#${existing.id})`,
    );
    return { id: String(existing.id), ...row };
  }
  const fields = {
    Title: row.candidateName,
    CandidateName: row.candidateName,
    CompanyNameLookupId: Number(row.companyId),
    Dateofbirth: `${row.dob}T00:00:00Z`,
    Status: "Active",
  };
  if (row.department) {
    fields.Department = row.department;
  }
  // Trainingmanager / Supervisor are lookup fields — skip unless IDs are known.
  try {
    const created = await createItem(client, LISTS.workforce, fields);
    console.log(`  workforce created: ${row.candidateName} (#${created.id})`);
    return { id: created.id, ...row };
  } catch (error) {
    // Retry without Department if choice value is not configured.
    if (row.department) {
      console.warn(
        `  workforce create with Department=${row.department} failed; retrying without it:`,
        error?.message ?? error,
      );
      delete fields.Department;
      const created = await createItem(client, LISTS.workforce, fields);
      console.log(
        `  workforce created (no dept): ${row.candidateName} (#${created.id})`,
      );
      return { id: created.id, ...row };
    }
    throw error;
  }
}

async function findPermission(client, email) {
  const items = await listItems(
    client,
    LISTS.permissions,
    `fields/UserEmail eq '${escapeOData(email)}'`,
    50,
  );
  return items[0] ?? null;
}

async function ensurePermission(client, row) {
  const existing = await findPermission(client, row.userEmail);
  const fields = {
    Title: row.userEmail,
    UserEmail: row.userEmail,
    RoleType: row.roleType,
    Status: "Active",
    AccessScope: row.accessScope,
    CanView: true,
    CanDownload: row.canDownload,
    CanEdit: row.canEdit,
  };
  if (row.companyId) {
    fields.CompanyLookupId = Number(row.companyId);
  }
  if (row.name) {
    fields.Name = row.name;
  }
  if (row.departments?.length) {
    fields.Departments = row.departments;
  }

  if (existing) {
    await updateItemFields(client, LISTS.permissions, existing.id, fields);
    console.log(`  permission updated: ${row.userEmail} (${row.roleType})`);
    return String(existing.id);
  }
  const created = await createItem(client, LISTS.permissions, fields);
  console.log(`  permission created: ${row.userEmail} (${row.roleType})`);
  return created.id;
}

async function findByCandidateLookup(client, listId, candidateId) {
  try {
    const filtered = await listItems(
      client,
      listId,
      `fields/CandidateNameLookupId eq ${Number(candidateId)}`,
      20,
    );
    if (filtered[0]) return filtered[0];
  } catch {
    // Fall through to light scan
  }
  const all = await listItems(client, listId, undefined, 200);
  return (
    all.find(
      (item) =>
        String(item.fields?.CandidateNameLookupId ?? "").trim() ===
        String(candidateId),
    ) ?? null
  );
}

async function ensureMatrix(client, row) {
  const existing = await findByCandidateLookup(
    client,
    LISTS.trainingMatrix,
    row.candidateId,
  );
  const fields = {
    Title: row.candidateName,
    CandidateNameLookupId: Number(row.candidateId),
    Company_x0020_NameLookupId: Number(row.companyId),
    MatrixCompanyLookupId: Number(row.companyId),
    Department: row.department ?? null,
    OverallStatus: row.overallStatus,
    NeedsReview: row.needsReview,
    NextExpiryDate: row.nextExpiryDate,
  };
  if (existing) {
    await updateItemFields(client, LISTS.trainingMatrix, existing.id, fields);
    console.log(`  matrix updated: ${row.candidateName}`);
    return String(existing.id);
  }
  const created = await createItem(client, LISTS.trainingMatrix, fields);
  console.log(`  matrix created: ${row.candidateName} (#${created.id})`);
  return created.id;
}

async function ensureRegister(client, listId, label, row) {
  const existing = await findByCandidateLookup(client, listId, row.candidateId);
  const payload = {
    Title: row.candidateName,
    CandidateNameLookupId: Number(row.candidateId),
    CompanyNameLookupId: Number(row.companyId),
    TrainingOutcome: "Pass",
    CustomerVisible: true,
    TrainingAddress: row.trainingAddress ?? "Test training address",
  };
  if (label === "In-House") {
    payload.CourseDate = row.trainingDate;
  } else if (row.trainingDate) {
    payload.TrainingDate = row.trainingDate;
  }
  if (existing) {
    try {
      await updateItemFields(client, listId, existing.id, payload);
      console.log(`  ${label} updated: ${row.candidateName}`);
      return String(existing.id);
    } catch (error) {
      console.warn(
        `  ${label} update skipped (${row.candidateName}):`,
        error?.message ?? error,
      );
      return String(existing.id);
    }
  }
  try {
    const created = await createItem(client, listId, payload);
    console.log(`  ${label} created: ${row.candidateName} (#${created.id})`);
    return created.id;
  } catch (error) {
    console.warn(`  ${label} create failed:`, error?.message ?? error);
    const minimal = {
      Title: row.candidateName,
      CandidateNameLookupId: Number(row.candidateId),
      CompanyNameLookupId: Number(row.companyId),
      CustomerVisible: true,
      TrainingOutcome: "Pass",
    };
    const created = await createItem(client, listId, minimal);
    console.log(
      `  ${label} created (minimal): ${row.candidateName} (#${created.id})`,
    );
    return created.id;
  }
}

async function ensureNvq(client, row) {
  const existing = await findByCandidateLookup(client, LISTS.nvq, row.candidateId);
  const fields = {
    Title: row.nvqTitle,
    CandidateNameLookupId: Number(row.candidateId),
    NVQCompanyLookupId: Number(row.companyId),
    NvqTitle: row.nvqTitle,
    CompletedDate: row.completedDate ?? null,
    CustomerVisible: true,
  };
  if (existing) {
    try {
      await updateItemFields(client, LISTS.nvq, existing.id, fields);
      console.log(`  nvq updated: ${row.nvqTitle}`);
    } catch (error) {
      console.warn(`  nvq update skipped:`, error?.message ?? error);
    }
    return String(existing.id);
  }
  try {
    const created = await createItem(client, LISTS.nvq, {
      ...fields,
      Company_x0020_NameLookupId: Number(row.companyId),
    });
    console.log(`  nvq created: ${row.nvqTitle} (#${created.id})`);
    return created.id;
  } catch (error) {
    console.warn(`  nvq create failed, retrying lean:`, error?.message ?? error);
    const created = await createItem(client, LISTS.nvq, fields);
    console.log(`  nvq created (lean): ${row.nvqTitle} (#${created.id})`);
    return created.id;
  }
}

async function ensureEvent(client, row) {
  const all = await listItems(client, LISTS.events, undefined, 2000);
  const existing = all.find(
    (item) =>
      String(item.fields?.Title ?? "").trim().toLowerCase() ===
      row.title.trim().toLowerCase(),
  );
  const fields = {
    Title: row.title,
    EventCompanyLookupId: Number(row.companyId),
    Customer_x0020_Visible: true,
    EventDate: row.eventDate,
    EndDate: row.endDate,
    TrainingAddress: row.trainingAddress,
    // DoNotSync is a text column on this list
    DoNotSync: "Yes",
    SyncStatus: "Skipped",
  };
  if (existing) {
    await updateItemFields(client, LISTS.events, existing.id, fields);
    console.log(`  event updated: ${row.title}`);
    return String(existing.id);
  }
  try {
    const created = await createItem(client, LISTS.events, fields);
    console.log(`  event created: ${row.title} (#${created.id})`);
    return created.id;
  } catch (error) {
    console.warn(`  event create failed, retrying lean:`, error?.message ?? error);
    const created = await createItem(client, LISTS.events, {
      Title: row.title,
      EventCompanyLookupId: Number(row.companyId),
      Customer_x0020_Visible: true,
      EventDate: row.eventDate,
      EndDate: row.endDate,
      TrainingAddress: row.trainingAddress,
    });
    console.log(`  event created (lean): ${row.title} (#${created.id})`);
    return created.id;
  }
}

function minimalPdf(label) {
  // Tiny valid-enough PDF bytes for Graph upload testing.
  const content = `%PDF-1.1
1 0 obj<<>>endobj
2 0 obj<< /Length 44 >>stream
BT /F1 12 Tf 100 700 Td (${label}) Tj ET
endstream
endobj
3 0 obj<< /Type /Page /Parent 4 0 R /Contents 2 0 R >>endobj
4 0 obj<< /Type /Pages /Kids [3 0 R] /Count 1 >>endobj
5 0 obj<< /Type /Catalog /Pages 4 0 R >>endobj
xref
0 6
trailer<< /Root 5 0 R /Size 6 >>
startxref
0
%%EOF`;
  return Buffer.from(content, "utf8");
}

async function ensureDocument(client, row) {
  // Upload into the Customer Documents library drive, then patch metadata.
  const drive = await client
    .api(`${siteRoot()}/lists/${LISTS.documents}/drive`)
    .get();
  const driveId = drive.id;
  const fileName = row.fileName;

  let itemId;
  try {
    const uploaded = await client
      .api(`/drives/${driveId}/root:/${encodeURIComponent(fileName)}:/content`)
      .header("Content-Type", "application/pdf")
      .put(minimalPdf(fileName));
    // Resolve list item id from drive item
    const listItem = await client
      .api(`/drives/${driveId}/items/${uploaded.id}/listItem`)
      .get();
    itemId = String(listItem.id);
  } catch (error) {
    console.warn(
      `  document upload failed for ${fileName}:`,
      error?.message ?? error,
    );
    return null;
  }

  const fields = {
    Title: fileName.replace(/\.pdf$/i, ""),
    CompanyLookupId: Number(row.companyId),
    DocumentType: row.documentType,
    CustomerVisible: row.customerVisible,
    NotificationSent: false,
  };
  if (row.candidateId) {
    fields.CandidateLookupId = Number(row.candidateId);
  }
  await updateItemFields(client, LISTS.documents, itemId, fields);
  console.log(
    `  document ready: ${fileName} (#${itemId}) visible=${row.customerVisible}`,
  );
  return itemId;
}

function daysFromNow(days) {
  const d = new Date();
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function hoursFromNow(hours) {
  const d = new Date();
  d.setUTCHours(d.getUTCHours() + hours);
  return d.toISOString();
}

async function main() {
  for (const [key, id] of Object.entries(LISTS)) {
    if (!id) throw new Error(`Missing list id for ${key}`);
  }
  if (!SITE) throw new Error("Missing SHAREPOINT_SITE_ID");

  const client = getClient();
  const summary = {
    companies: {},
    workforce: {},
    permissions: {},
    documents: {},
    notes: [],
  };

  console.log("\n=== Step 1: Companies ===");
  const murphy = await ensureCompany(client, "Murphy plant", "TEST-MURPHY-001");
  const fast = await ensureCompany(
    client,
    "Fast - Zohaib Rashid",
    "TEST-FAST-001",
  );
  summary.companies.murphy = murphy;
  summary.companies.fast = fast;

  console.log("\n=== Step 2: Workforce candidates ===");
  const john = await ensureWorkforce(client, {
    candidateName: "John Murphy Test",
    companyId: murphy.id,
    companyName: murphy.name,
    department: "Groundworks",
    dob: "1990-01-01",
  });
  const ali = await ensureWorkforce(client, {
    candidateName: "Ali Murphy Test",
    companyId: murphy.id,
    companyName: murphy.name,
    department: "Plant",
    dob: "1992-02-02",
  });
  const fastCandidate = await ensureWorkforce(client, {
    candidateName: "Fast Candidate Test",
    companyId: fast.id,
    companyName: fast.name,
    department: "Admin",
    dob: "1993-03-03",
  });
  summary.notes.push(
    "Workforce Trainingmanager/Supervisor are lookup fields — set manually to manager/supervisor test users in SharePoint if needed.",
  );
  summary.workforce.john = john;
  summary.workforce.ali = ali;
  summary.workforce.fastCandidate = fastCandidate;

  console.log("\n=== Step 3: Permissions ===");
  summary.permissions.admin = await ensurePermission(client, {
    userEmail: "haseeb@pavetraining.co.uk",
    roleType: "Admin",
    accessScope: "All",
    canDownload: true,
    canEdit: true,
    companyId: null,
  });
  summary.permissions.manager = await ensurePermission(client, {
    userEmail: "manager-test@murphyplant.com",
    roleType: "Training Manager",
    accessScope: "Company",
    canDownload: true,
    canEdit: false,
    companyId: murphy.id,
  });
  try {
    summary.permissions.supervisor = await ensurePermission(client, {
      userEmail: "supervisor-test@murphyplant.com",
      roleType: "Supervisor",
      accessScope: "Department",
      canDownload: false,
      canEdit: false,
      companyId: murphy.id,
      name: "supervisor test user",
    });
    summary.notes.push(
      "Set Permissions.Departments = Groundworks manually for supervisor-test@murphyplant.com if choice values allow it.",
    );
  } catch (error) {
    console.warn("  supervisor permission failed:", error?.message ?? error);
  }
  try {
    summary.permissions.candidate = await ensurePermission(client, {
      userEmail: "candidate-test@email.com",
      roleType: "Candidate",
      accessScope: "CandidateOnly",
      canDownload: true,
      canEdit: false,
      companyId: murphy.id,
      name: "John Murphy Test",
    });
  } catch (error) {
    console.warn(
      "  Candidate role failed; writing Supervisor + CandidateOnly:",
      error?.message ?? error,
    );
    summary.permissions.candidate = await ensurePermission(client, {
      userEmail: "candidate-test@email.com",
      roleType: "Supervisor",
      accessScope: "CandidateOnly",
      canDownload: true,
      canEdit: false,
      companyId: murphy.id,
      name: "John Murphy Test",
    });
    summary.notes.push(
      "RoleType=Candidate may not exist yet — used Supervisor + CandidateOnly + Name=John Murphy Test.",
    );
  }

  console.log("\n=== Step 4: Documents ===");
  summary.documents.murphyVisible = await ensureDocument(client, {
    fileName: "john-murphy-certificate.pdf",
    companyId: murphy.id,
    candidateId: john.id,
    documentType: "Certificate",
    customerVisible: true,
  });
  summary.documents.murphyHidden = await ensureDocument(client, {
    fileName: "murphy-internal-note.pdf",
    companyId: murphy.id,
    candidateId: john.id,
    documentType: "Internal",
    customerVisible: false,
  });
  summary.documents.fastVisible = await ensureDocument(client, {
    fileName: "fast-certificate.pdf",
    companyId: fast.id,
    candidateId: fastCandidate.id,
    documentType: "Certificate",
    customerVisible: true,
  });

  console.log("\n=== Step 5: Training Matrix ===");
  await ensureMatrix(client, {
    candidateName: john.candidateName,
    candidateId: john.id,
    companyId: murphy.id,
    companyName: murphy.name,
    department: "Groundworks",
    overallStatus: "Expired",
    needsReview: true,
    nextExpiryDate: daysFromNow(-30),
  });
  await ensureMatrix(client, {
    candidateName: ali.candidateName,
    candidateId: ali.id,
    companyId: murphy.id,
    companyName: murphy.name,
    department: "Plant",
    overallStatus: "Valid",
    needsReview: false,
    nextExpiryDate: daysFromNow(300),
  });
  await ensureMatrix(client, {
    candidateName: fastCandidate.candidateName,
    candidateId: fastCandidate.id,
    companyId: fast.id,
    companyName: fast.name,
    department: "Admin",
    overallStatus: "Valid",
    needsReview: false,
    nextExpiryDate: daysFromNow(180),
  });

  console.log("\n=== Step 6: Training registers ===");
  const registerPairs = [
    { listId: LISTS.npors, label: "NPORS" },
    { listId: LISTS.eusr, label: "EUSR" },
    { listId: LISTS.nrswa, label: "Streetworks" },
    { listId: LISTS.inHouse, label: "In-House" },
  ];

  for (const reg of registerPairs) {
    await ensureRegister(client, reg.listId, reg.label, {
      candidateName: john.candidateName,
      candidateId: john.id,
      companyId: murphy.id,
      trainingDate: daysFromNow(-60),
      trainingAddress: "Test training address",
    });
    await ensureRegister(client, reg.listId, reg.label, {
      candidateName: fastCandidate.candidateName,
      candidateId: fastCandidate.id,
      companyId: fast.id,
      trainingDate: daysFromNow(-45),
      trainingAddress: "Fast test training address",
    });
  }

  console.log("\n=== Step 7: NVQ ===");
  await ensureNvq(client, {
    candidateName: john.candidateName,
    candidateId: john.id,
    companyId: murphy.id,
    companyName: murphy.name,
    nvqTitle: "Test NVQ",
    stageOfNvq: "In Progress",
    completedDate: null,
  });
  await ensureNvq(client, {
    candidateName: ali.candidateName,
    candidateId: ali.id,
    companyId: murphy.id,
    companyName: murphy.name,
    nvqTitle: "Completed Test NVQ",
    stageOfNvq: "Completed",
    completedDate: daysFromNow(0),
  });

  console.log("\n=== Step 8: Events ===");
  await ensureEvent(client, {
    title: "Murphy Test Booking",
    companyId: murphy.id,
    eventDate: hoursFromNow(24 * 7),
    endDate: hoursFromNow(24 * 7 + 2),
    trainingAddress: "Test training address",
  });
  await ensureEvent(client, {
    title: "Fast Test Booking",
    companyId: fast.id,
    eventDate: hoursFromNow(24 * 8),
    endDate: hoursFromNow(24 * 8 + 2),
    trainingAddress: "Fast training address",
  });

  console.log("\n=== Seed complete ===");
  console.log(
    JSON.stringify(
      {
        murphyCompanyId: murphy.id,
        fastCompanyId: fast.id,
        johnWorkforceId: john.id,
        aliWorkforceId: ali.id,
        fastWorkforceId: fastCandidate.id,
        documents: summary.documents,
        notes: summary.notes,
        loginHints: {
          admin: "haseeb@pavetraining.co.uk → /admin",
          trainingManager:
            "manager-test@murphyplant.com (must exist in Entra to login)",
          supervisor:
            "supervisor-test@murphyplant.com (must exist in Entra to login)",
          candidate:
            "candidate-test@email.com (must exist in Entra to login)",
        },
      },
      null,
      2,
    ),
  );
}

main().catch((error) => {
  console.error("\nSeed failed:", error?.message ?? error);
  if (error?.body) console.error(error.body);
  process.exit(1);
});
