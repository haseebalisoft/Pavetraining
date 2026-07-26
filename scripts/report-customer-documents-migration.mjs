/**
 * Read-only Customer Documents folder migration report.
 * Does NOT move or rename files — prints mismatches for review.
 *
 * Usage (from repo root):
 *   node --env-file=.env.local scripts/report-customer-documents-migration.mjs
 */

import { writeFileSync } from "node:fs";
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const DOCUMENTS_LIST_ID = process.env.SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID;
const COMPANY_LIST_ID = process.env.SHAREPOINT_COMPANY_LIST_ID;
const WORKFORCE_LIST_ID = process.env.SHAREPOINT_WORKFORCE_LIST_ID;
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
  if (siteId.includes(":/")) {
    const withTransition = siteId.endsWith(":") ? siteId : `${siteId}:`;
    return `/sites/${withTransition}`;
  }
  return `/sites/${siteId}`;
}

function sanitizeFolderSegment(value) {
  return String(value ?? "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

function companyFolderName(companyNumber, companyName) {
  const name = sanitizeFolderSegment(companyName) || "Company";
  const number = sanitizeFolderSegment(companyNumber ?? "");
  return number ? `${number} - ${name}` : name;
}

function candidateFolderName(workforceNumber, candidateName) {
  const name = sanitizeFolderSegment(candidateName) || "Candidate";
  const number = sanitizeFolderSegment(workforceNumber ?? "");
  return number ? `${number} - ${name}` : name;
}

function resolveDestinationFolder(documentType, hasCandidate) {
  if (!hasCandidate) return "Company Documents";
  const normalized = String(documentType ?? "")
    .trim()
    .toLowerCase();
  if (normalized.includes("certificate")) return "Certificates";
  if (
    normalized.includes("card scan") ||
    (normalized.includes("card") && normalized.includes("scan"))
  ) {
    return "Card Scans";
  }
  if (normalized.includes("nvq")) return "NVQ Documents";
  return "Other Documents";
}

function expectedSegments(input) {
  const destination = resolveDestinationFolder(
    input.documentType,
    input.hasCandidate,
  );
  const company = companyFolderName(input.companyNumber, input.companyName);
  if (!input.hasCandidate || destination === "Company Documents") {
    return {
      segments: [company, "Company Documents"],
      destinationFolder: "Company Documents",
    };
  }
  return {
    segments: [
      company,
      "Candidates",
      candidateFolderName(input.workforceNumber, input.candidateName),
      destination,
    ],
    destinationFolder: destination,
  };
}

function decodeSegment(segment) {
  try {
    return decodeURIComponent(String(segment).replace(/\+/g, " "));
  } catch {
    return String(segment);
  }
}

function extractFolderPath(fileRef) {
  if (!fileRef) return [];
  const parts = String(fileRef)
    .split("/")
    .map(decodeSegment)
    .filter(Boolean);
  const idx = parts.findIndex(
    (part) => part.toLowerCase() === "customer documents",
  );
  const after = idx >= 0 ? parts.slice(idx + 1) : parts;
  if (after.length <= 1) return [];
  return after.slice(0, -1);
}

function pathsMatch(actual, expected, companyNumber, workforceNumber) {
  if (actual.length !== expected.length) return false;
  return actual.every((segment, index) => {
    const want = expected[index];
    if (segment.trim().toLowerCase() === want.trim().toLowerCase()) {
      return true;
    }
    if (index === 0 && companyNumber) {
      const prefix = `${String(companyNumber).trim().toLowerCase()} -`;
      const lower = segment.toLowerCase();
      return (
        lower.startsWith(prefix) ||
        lower === String(companyNumber).trim().toLowerCase()
      );
    }
    if (
      index === 2 &&
      workforceNumber &&
      expected[1]?.toLowerCase() === "candidates"
    ) {
      const prefix = `${String(workforceNumber).trim().toLowerCase()} -`;
      const lower = segment.toLowerCase();
      return (
        lower.startsWith(prefix) ||
        lower === String(workforceNumber).trim().toLowerCase()
      );
    }
    return false;
  });
}

async function listAllItems(client, listId) {
  const items = [];
  let url = `${siteRoot()}/lists/${listId}/items?$expand=fields&$top=200`;
  while (url) {
    const response = await client
      .api(url.replace("https://graph.microsoft.com/v1.0", ""))
      .header("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly")
      .get();
    items.push(...(response.value ?? []));
    url = response["@odata.nextLink"] ?? null;
  }
  return items;
}

function field(item, name) {
  return item.fields?.[name];
}

function asString(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === "object" && value !== null) {
    const lookup =
      value.LookupValue ?? value.Title ?? value.DisplayName ?? value.Label;
    if (lookup != null) return String(lookup).trim() || null;
    if (value.LookupId != null) return String(value.LookupId);
  }
  const text = String(value).trim();
  return text || null;
}

async function main() {
  requireEnv("SHAREPOINT_SITE_ID");
  if (!DOCUMENTS_LIST_ID) throw new Error("Missing SHAREPOINT_CUSTOMER_DOCUMENTS_LIST_ID");
  if (!COMPANY_LIST_ID) throw new Error("Missing SHAREPOINT_COMPANY_LIST_ID");
  if (!WORKFORCE_LIST_ID) throw new Error("Missing SHAREPOINT_WORKFORCE_LIST_ID");

  const client = getClient();
  const [docItems, companyItems, workforceItems] = await Promise.all([
    listAllItems(client, DOCUMENTS_LIST_ID),
    listAllItems(client, COMPANY_LIST_ID),
    listAllItems(client, WORKFORCE_LIST_ID),
  ]);

  const companies = new Map();
  for (const item of companyItems) {
    companies.set(String(item.id), {
      id: String(item.id),
      companyName:
        asString(field(item, "CompanyName")) ??
        asString(field(item, "Title")) ??
        "",
      companyNumber: asString(field(item, "CompanyNumber")),
    });
  }

  const workforce = new Map();
  for (const item of workforceItems) {
    workforce.set(String(item.id), {
      id: String(item.id),
      candidateName: asString(field(item, "CandidateName")) ?? "",
      workforceNumber: asString(field(item, "WorkforceNumber")),
      companyName: asString(field(item, "CompanyName")),
    });
  }

  const mismatches = [];
  let scannedFiles = 0;
  let matched = 0;
  let skipped = 0;

  for (const item of docItems) {
    const fs = field(item, "FSObjType");
    if (fs === 1 || fs === "1") {
      skipped += 1;
      continue;
    }
    const fileName =
      asString(field(item, "FileLeafRef")) ?? asString(field(item, "Title"));
    if (!fileName) {
      skipped += 1;
      continue;
    }
    scannedFiles += 1;

    const companyId = asString(field(item, "CompanyLookupId"));
    const candidateId = asString(field(item, "CandidateLookupId"));
    const documentType = asString(field(item, "DocumentType"));
    const fileRef = asString(field(item, "FileRef"));
    const current = extractFolderPath(fileRef);

    const company = companyId ? companies.get(companyId) : null;
    const companyName =
      company?.companyName ||
      asString(field(item, "Company")) ||
      null;
    if (!companyName) {
      mismatches.push({
        id: String(item.id),
        fileName,
        reason: "Missing company metadata",
        currentFolderPath: current.join("/"),
        expectedFolderPath: "",
        fileRef,
      });
      continue;
    }

    const candidate = candidateId ? workforce.get(candidateId) : null;
    const hasCandidate = Boolean(candidateId);
    const { segments, destinationFolder } = expectedSegments({
      companyNumber: company?.companyNumber ?? null,
      companyName,
      workforceNumber: candidate?.workforceNumber ?? null,
      candidateName:
        candidate?.candidateName || asString(field(item, "Candidate")) || "Candidate",
      documentType,
      hasCandidate,
    });

    if (
      pathsMatch(
        current,
        segments,
        company?.companyNumber ?? null,
        candidate?.workforceNumber ?? null,
      )
    ) {
      matched += 1;
      continue;
    }

    mismatches.push({
      id: String(item.id),
      fileName,
      company: companyName,
      companyId,
      candidate: candidate?.candidateName ?? asString(field(item, "Candidate")),
      candidateId,
      documentType,
      destinationFolder,
      currentFolderPath: current.join("/"),
      expectedFolderPath: segments.join("/"),
      fileRef,
      reason: current.length
        ? "Path does not match expected structure"
        : "No folder path under Customer Documents",
    });
  }

  const report = {
    generatedAt: new Date().toISOString(),
    scannedFiles,
    matched,
    mismatched: mismatches.length,
    skipped,
    notes: [
      "Read-only report — no files were moved.",
      "Review URL/integration impact before any migration.",
    ],
    mismatches,
  };

  const outPath = "customer-documents-migration-report.json";
  writeFileSync(outPath, JSON.stringify(report, null, 2), "utf8");

  console.log(
    `Scanned ${scannedFiles} file(s): ${matched} matched, ${mismatches.length} mismatch(es), ${skipped} skipped.`,
  );
  console.log(`Wrote ${outPath}`);
  console.log("No files were moved.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
