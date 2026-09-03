import "server-only";

import {
  createAdminNvq,
  createAdminRegister,
  listAdminCompanies,
  listAdminNvq,
  listAdminRegister,
  listAdminWorkforce,
  updateAdminNvq,
  updateAdminRegister,
  type AdminNvqRecord,
  type AdminRegisterKey,
  type AdminTrainingRecord,
  type AdminWorkforceRecord,
} from "@/lib/services/adminCrudService";
import { summarizeBulkRows } from "@/lib/services/bulkUpload/candidateImporter";
import {
  findCompanyByName,
  nameKey,
  sameDate,
} from "@/lib/services/bulkUpload/matching";
import {
  normalizeDateValue,
  pickField,
  type ParsedSpreadsheet,
} from "@/lib/services/bulkUpload/parseSpreadsheet";
import { triggerMatrixSyncAfterRegister } from "@/lib/services/matrixSyncHook";
import type {
  BulkCommitRowInput,
  BulkDuplicateMode,
  BulkImportType,
  BulkPreviewRow,
} from "@/types/bulkUpload";
import type { Company } from "@/types/models";

export type RegisterImportType =
  | "npors"
  | "eusr"
  | "streetworks"
  | "inHouse"
  | "nvq";

const REGISTER_KEY_BY_TYPE: Record<
  Exclude<RegisterImportType, "nvq">,
  AdminRegisterKey
> = {
  npors: "nporsRegister",
  eusr: "eusrRegister",
  streetworks: "nrswaRegister",
  inHouse: "inHouseCertificates",
};

const NAME_ALIASES = [
  "Candidate Name",
  "CandidateName",
  "Name",
  "Full Name",
];
const COMPANY_ALIASES = [
  "Company",
  "Company Name",
  "CompanyName",
  "NVQ Company",
];

function isRegisterImportType(
  value: BulkImportType,
): value is RegisterImportType {
  return (
    value === "npors" ||
    value === "eusr" ||
    value === "streetworks" ||
    value === "inHouse" ||
    value === "nvq"
  );
}

function mapCommonFields(
  raw: Record<string, string | null>,
): Record<string, string | null> {
  return {
    candidateName: pickField(raw, NAME_ALIASES),
    company: pickField(raw, COMPANY_ALIASES),
    companyName: pickField(raw, COMPANY_ALIASES),
    trainingDate: normalizeDateValue(
      pickField(raw, [
        "Start Date",
        "Training Date",
        "TrainingDate",
        "Course Date",
        "CourseDate",
        "Date Registered",
        "DateRegistered",
      ]),
    ),
    expiry: normalizeDateValue(
      pickField(raw, ["Expiry", "Expiry Date", "ExpiryDate", "Expirydate"]),
    ),
    trainingOutcome: pickField(raw, [
      "Outcome",
      "Training Outcome",
      "TrainingOutcome",
    ]),
    trainingAddress: pickField(raw, [
      "Training Address",
      "TrainingAddress",
      "Address",
    ]),
    notes: pickField(raw, ["Notes", "Customer Update Notes", "Outcome Notes"]),
  };
}

function mapRegisterSpecificFields(
  importType: RegisterImportType,
  raw: Record<string, string | null>,
): Record<string, string | null> {
  const common = mapCommonFields(raw);

  if (importType === "npors") {
    return {
      ...common,
      nporsNumber: pickField(raw, [
        "NPORS Number",
        "NPORSNumber",
        "On/Number",
        "On Number",
      ]),
      nporsCategory: pickField(raw, [
        "NPORS Category",
        "NPORSCategory",
        "Category",
      ]),
      noviceOrEwt: pickField(raw, [
        "Novice or EWT",
        "Novice or Ewt",
        "NoviceorEwt",
        "Novice/EWT",
      ]),
    };
  }

  if (importType === "eusr") {
    return {
      ...common,
      eusrNumber: pickField(raw, ["EUSR Number", "EUSRNumber"]),
      eusrCategory: pickField(raw, [
        "EUSR Category",
        "Eusr Category",
        "EusrCategory",
        "Category",
      ]),
      cardType: pickField(raw, ["Card Type", "CardType"]),
    };
  }

  if (importType === "streetworks") {
    return {
      ...common,
      swqrNumber: pickField(raw, ["SWQR Number", "SWQRNumber"]),
      course: pickField(raw, ["Course"]),
      streetworksCategory: pickField(raw, [
        "Streetworks Category",
        "StreetworksCategory",
        "NRSWA Category",
        "Category",
      ]),
    };
  }

  if (importType === "inHouse") {
    return {
      ...common,
      course: pickField(raw, ["Course", "Course Category", "CourseCategory"]),
      courseCategory: pickField(raw, [
        "Course",
        "Course Category",
        "CourseCategory",
      ]),
      certificateCategory: pickField(raw, [
        "Certificate Category",
        "CertificateCategory",
      ]),
      inHouseCertificationNumber: pickField(raw, [
        "Certification Number",
        "In-House Certification Number",
        "In House Certification Number",
        "CertificationNumber",
      ]),
    };
  }

  return {
    ...common,
    nvqTitle: pickField(raw, ["NVQ Title", "Nvq Title", "NvqTitle", "Title"]),
    boltOn: pickField(raw, [
      "Bolt-on NVQ",
      "Bolt On",
      "Bolt on Nvq",
      "Bolt On NVQ",
      "BoltonNvq",
      "BoltOn",
    ]),
    niNumber: pickField(raw, ["NI Number", "NINumber", "Ni Number"]),
    ulnNumber: pickField(raw, ["ULN Number", "ULNNumber", "ULN"]),
    poNumber: pickField(raw, ["PO Number", "PONumber", "PO"]),
    cardSchemeCategory: pickField(raw, [
      "Card Scheme Category",
      "CardSchemeCategory",
    ]),
    cardExtensionDateNeeded: normalizeDateValue(
      pickField(raw, [
        "Card Extension Date Needed",
        " Card Extension Date Needed",
        "CardExtensionDateNeeded",
      ]),
    ),
    siteAddress: pickField(raw, ["Site Address", "SiteAddress"]),
    siteContact: pickField(raw, [
      "Site Contact Name/Number",
      "Site Contact Name/Number",
      "SiteContactNameNumber",
      "Site Contact",
    ]),
    englishUnderstandingConfirmed: pickField(raw, [
      "English Understanding Confirmed",
      "EnglishUnderstandingConfirmed",
    ]),
    tcAcknowledged: pickField(raw, [
      "T&C Acknowledged",
      "TC Acknowledged",
      "TCAcknowledged",
    ]),
    gdprConsent: pickField(raw, ["GDPR Consent", "GDPRConsent"]),
    dateRegistered: normalizeDateValue(
      pickField(raw, [
        "Date Registered",
        "DateRegistered",
        "Start Date",
        "Registered",
      ]),
    ),
    inductionDate: normalizeDateValue(
      pickField(raw, [
        "Date Induction Booked",
        "Date induction booked",
        "DateinductionBooked",
        "Induction Date",
      ]),
    ),
    stageOfNvq: pickField(raw, [
      "Stage of NVQ",
      "Stage of Nvq",
      "StageofNvq",
      "Stage",
    ]),
    notes: pickField(raw, ["Notes"]),
    customerUpdateNotes: pickField(raw, [
      "Customer Update Notes",
      " Customer Update Notes",
      "CustomerUpdateNotes",
    ]),
    completedDate: normalizeDateValue(
      pickField(raw, ["Completed Date", "CompletedDate"]),
    ),
    certificationDate: normalizeDateValue(
      pickField(raw, ["Certification Date", "CertificationDate"]),
    ),
    trainingOutcome: pickField(raw, [
      "TrainingOutcome",
      "Training Outcome",
      "Outcome",
    ]),
    outcomeDate: normalizeDateValue(
      pickField(raw, ["OutcomeDate", "Outcome Date"]),
    ),
    assessorTrainer: pickField(raw, [
      "AssessorTrainer",
      "Assessor Trainer",
      "Assessor",
    ]),
    outcomeNotes: pickField(raw, ["OutcomeNotes", "Outcome Notes"]),
    customerVisible: pickField(raw, ["CustomerVisible", "Customer Visible"]),
  };
}

function findWorkforceMatch(
  workforce: AdminWorkforceRecord[],
  candidateName: string,
  companyName: string,
): AdminWorkforceRecord | null {
  const cName = nameKey(candidateName);
  const companyExact = nameKey(companyName);
  // Company-scoped match only — a same-named candidate at a different
  // company must surface as "not found for that company" (existing message
  // path), never silently attach the training record to the wrong company.
  return (
    workforce.find(
      (row) =>
        nameKey(row.candidateName) === cName &&
        nameKey(row.companyName) === companyExact,
    ) ?? null
  );
}

function findRegisterDuplicate(
  importType: Exclude<RegisterImportType, "nvq">,
  existing: AdminTrainingRecord[],
  fields: Record<string, string | null>,
): AdminTrainingRecord | null {
  const candidate = nameKey(fields.candidateName);
  const company = nameKey(fields.company ?? fields.companyName);
  const trainingDate = fields.trainingDate;
  const expiry = fields.expiry;

  return (
    existing.find((row) => {
      if (nameKey(row.candidateName) !== candidate) return false;
      if (nameKey(row.companyName) !== company) return false;

      if (importType === "npors") {
        const cat = nameKey(fields.nporsCategory);
        if (cat && nameKey(row.nporsCategory) === cat) {
          return (
            sameDate(row.expiry, expiry) ||
            sameDate(row.trainingDate, trainingDate) ||
            (!expiry && !trainingDate)
          );
        }
        if (
          fields.nporsNumber &&
          nameKey(row.nporsNumber) === nameKey(fields.nporsNumber)
        ) {
          return (
            sameDate(row.expiry, expiry) ||
            sameDate(row.trainingDate, trainingDate)
          );
        }
      }

      if (importType === "eusr") {
        const cat = nameKey(fields.eusrCategory);
        if (cat && nameKey(row.eusrCategory) === cat) {
          return (
            sameDate(row.expiry, expiry) ||
            sameDate(row.trainingDate, trainingDate)
          );
        }
      }

      if (importType === "streetworks") {
        const cat = nameKey(fields.streetworksCategory ?? fields.course);
        const rowCat = nameKey(row.streetworksCategory ?? row.course);
        if (cat && cat === rowCat) {
          return (
            sameDate(row.expiry, expiry) ||
            sameDate(row.trainingDate, trainingDate)
          );
        }
      }

      if (importType === "inHouse") {
        const course = nameKey(fields.course ?? fields.courseCategory);
        const rowCourse = nameKey(row.course ?? row.courseCategory);
        if (course && course === rowCourse) {
          return (
            sameDate(row.expiry, expiry) ||
            sameDate(row.trainingDate, trainingDate)
          );
        }
      }

      return (
        Boolean(trainingDate) &&
        sameDate(row.trainingDate, trainingDate) &&
        sameDate(row.expiry, expiry)
      );
    }) ?? null
  );
}

function findNvqDuplicate(
  existing: AdminNvqRecord[],
  fields: Record<string, string | null>,
): AdminNvqRecord | null {
  const candidate = nameKey(fields.candidateName);
  const company = nameKey(fields.company ?? fields.companyName);
  const title = nameKey(fields.nvqTitle);
  return (
    existing.find((row) => {
      if (nameKey(row.candidateName) !== candidate) return false;
      if (company && nameKey(row.companyName) !== company) return false;
      if (title && nameKey(row.nvqTitle) === title) return true;
      return (
        Boolean(fields.dateRegistered) &&
        sameDate(row.dateRegistered, fields.dateRegistered)
      );
    }) ?? null
  );
}

function toCreateInput(
  importType: RegisterImportType,
  fields: Record<string, string | null>,
  resolved: { candidateName: string; companyName: string; workforceId: string },
): Record<string, unknown> {
  if (importType === "nvq") {
    return {
      workforceId: resolved.workforceId,
      candidateName: resolved.candidateName,
      companyName: resolved.companyName,
      niNumber: fields.niNumber,
      ulnNumber: fields.ulnNumber,
      nvqTitle: fields.nvqTitle,
      boltOn: fields.boltOn,
      poNumber: fields.poNumber,
      cardSchemeCategory: fields.cardSchemeCategory,
      cardExtensionDateNeeded: fields.cardExtensionDateNeeded,
      siteAddress: fields.siteAddress,
      siteContact: fields.siteContact,
      englishUnderstandingConfirmed: fields.englishUnderstandingConfirmed,
      tcAcknowledged: fields.tcAcknowledged,
      gdprConsent: fields.gdprConsent,
      dateRegistered: fields.dateRegistered ?? fields.trainingDate,
      inductionDate: fields.inductionDate,
      stageOfNvq: fields.stageOfNvq,
      notes: fields.notes,
      customerUpdateNotes: fields.customerUpdateNotes,
      completedDate: fields.completedDate,
      certificationDate: fields.certificationDate,
      trainingOutcome: fields.trainingOutcome,
      outcomeDate: fields.outcomeDate,
      assessorTrainer: fields.assessorTrainer,
      outcomeNotes: fields.outcomeNotes,
      customerVisible: fields.customerVisible ?? true,
    };
  }

  return {
    workforceId: resolved.workforceId,
    candidateName: resolved.candidateName,
    companyName: resolved.companyName,
    trainingDate: fields.trainingDate,
    trainingAddress: fields.trainingAddress,
    trainingOutcome: fields.trainingOutcome,
    expiry: fields.expiry,
    notes: fields.notes,
    nporsNumber: fields.nporsNumber,
    nporsCategory: fields.nporsCategory,
    noviceOrEwt: fields.noviceOrEwt,
    eusrCategory: fields.eusrCategory,
    cardType: fields.cardType,
    course: fields.course ?? fields.courseCategory,
    courseCategory: fields.courseCategory ?? fields.course,
    streetworksCategory: fields.streetworksCategory,
    certificateCategory:
      fields.certificateCategory ?? fields.course ?? fields.courseCategory,
    inHouseCertificationNumber: fields.inHouseCertificationNumber,
    customerVisible: true,
  };
}

async function buildPreviewRows(
  importType: RegisterImportType,
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  const [companies, workforce, existingRegisters, existingNvq] =
    await Promise.all([
      listAdminCompanies(),
      listAdminWorkforce(),
      importType === "nvq"
        ? Promise.resolve([] as AdminTrainingRecord[])
        : listAdminRegister(REGISTER_KEY_BY_TYPE[importType]),
      importType === "nvq"
        ? listAdminNvq()
        : Promise.resolve([] as AdminNvqRecord[]),
    ]);

  return spreadsheet.rows.map((raw, index) => {
    const fields = mapRegisterSpecificFields(importType, raw);
    const messages: string[] = [];
    const candidateName = fields.candidateName?.trim() || "";
    const companyName = (fields.company ?? fields.companyName)?.trim() || "";

    if (!candidateName) messages.push("Candidate Name is required.");
    if (!companyName) messages.push("Company is required.");

    let company: Company | null = null;
    let candidate: AdminWorkforceRecord | null = null;

    if (companyName) {
      company = findCompanyByName(companies, companyName);
      if (!company) {
        messages.push(
          `Company "${companyName}" was not found. Import Companies / Workforce first.`,
        );
      }
    }

    if (candidateName && company) {
      candidate = findWorkforceMatch(
        workforce,
        candidateName,
        company.companyName,
      );
      if (!candidate) {
        messages.push(
          `Candidate "${candidateName}" was not found in Workforce for that company.`,
        );
      }
    }

    if (
      fields.trainingOutcome &&
      !/^(pass|fail)$/i.test(fields.trainingOutcome.trim())
    ) {
      messages.push("Outcome must be Pass or Fail when provided.");
    }

    let matchedEntityId: string | null = null;
    let matchedEntityName: string | null = null;
    let duplicateMatch: BulkPreviewRow["duplicateMatch"] = null;

    if (candidate && company && messages.length === 0) {
      if (importType === "nvq") {
        const dup = findNvqDuplicate(existingNvq, {
          ...fields,
          candidateName: candidate.candidateName,
          company: company.companyName,
        });
        if (dup) {
          matchedEntityId = dup.id;
          matchedEntityName = `${dup.candidateName} · ${dup.nvqTitle ?? "NVQ"}`;
          duplicateMatch = "nameCompany";
          messages.push(
            "Matching NVQ row found — will update or skip based on mode.",
          );
        }
      } else {
        const dup = findRegisterDuplicate(importType, existingRegisters, {
          ...fields,
          candidateName: candidate.candidateName,
          company: company.companyName,
        });
        if (dup) {
          matchedEntityId = dup.id;
          matchedEntityName = `${dup.candidateName} · ${dup.trainingDate ?? dup.expiry ?? dup.id}`;
          duplicateMatch = "nameCompany";
          messages.push(
            "Matching register row found — will update or skip based on mode.",
          );
        }
      }
    }

    let status: BulkPreviewRow["status"] = "Ready";
    if (messages.some((m) => /required|not found|must be/i.test(m))) {
      status = "Error";
    } else if (matchedEntityId) {
      status = "Duplicate";
    } else if (messages.length) {
      status = "Warning";
    }

    return {
      rowNumber: index + 2,
      status,
      messages,
      fields: {
        ...fields,
        candidateName: candidate?.candidateName ?? (candidateName || null),
        company: company?.companyName ?? (companyName || null),
        companyName: company?.companyName ?? (companyName || null),
        workforceId: candidate?.id ?? null,
        matchedEntityId,
      },
      source: raw,
      resolvedCompanyName: company?.companyName ?? null,
      matchedEntityId,
      matchedEntityName,
      duplicateMatch,
    };
  });
}

export async function previewRegisterImport(
  importType: BulkImportType,
  spreadsheet: ParsedSpreadsheet,
): Promise<BulkPreviewRow[]> {
  if (!isRegisterImportType(importType)) {
    throw new Error(`Unsupported register import type: ${importType}`);
  }
  return buildPreviewRows(importType, spreadsheet);
}

export async function commitRegisterImport(input: {
  importType: BulkImportType;
  rows: BulkCommitRowInput[];
  duplicateMode: BulkDuplicateMode;
}): Promise<BulkPreviewRow[]> {
  if (!isRegisterImportType(input.importType)) {
    throw new Error(`Unsupported register import type: ${input.importType}`);
  }
  const importType = input.importType;
  const [companies, workforce] = await Promise.all([
    listAdminCompanies(),
    listAdminWorkforce(),
  ]);

  const results: BulkPreviewRow[] = [];

  for (const row of input.rows) {
    const fields = { ...row.fields };
    const candidateName = fields.candidateName?.trim() || "";
    const companyName = (fields.company ?? fields.companyName)?.trim() || "";

    if (!candidateName || !companyName) {
      results.push({
        rowNumber: row.rowNumber,
        status: "Error",
        messages: ["Candidate Name and Company are required."],
        fields,
        matchedEntityId: null,
        matchedEntityName: null,
        duplicateMatch: null,
      });
      continue;
    }

    const company = findCompanyByName(companies, companyName);
    const candidate = company
      ? findWorkforceMatch(workforce, candidateName, company.companyName)
      : null;

    if (!company || !candidate) {
      results.push({
        rowNumber: row.rowNumber,
        status: "Error",
        messages: [
          !company
            ? `Company "${companyName}" was not found.`
            : `Candidate "${candidateName}" was not found in Workforce.`,
        ],
        fields,
        matchedEntityId: null,
        matchedEntityName: null,
        duplicateMatch: null,
      });
      continue;
    }

    const resolved = {
      candidateName: candidate.candidateName,
      companyName: company.companyName,
      workforceId: candidate.id,
    };
    const payload = toCreateInput(importType, fields, resolved);

    try {
      let existingId = fields.matchedEntityId?.trim() || null;
      if (!existingId) {
        if (importType === "nvq") {
          const dup = findNvqDuplicate(await listAdminNvq(), {
            ...fields,
            candidateName: resolved.candidateName,
            company: resolved.companyName,
          });
          existingId = dup?.id ?? null;
        } else {
          const dup = findRegisterDuplicate(
            importType,
            await listAdminRegister(REGISTER_KEY_BY_TYPE[importType]),
            {
              ...fields,
              candidateName: resolved.candidateName,
              company: resolved.companyName,
            },
          );
          existingId = dup?.id ?? null;
        }
      }

      if (existingId && input.duplicateMode === "skip") {
        results.push({
          rowNumber: row.rowNumber,
          status: "Skipped",
          messages: ["Duplicate skipped."],
          fields: {
            ...fields,
            candidateName: resolved.candidateName,
            company: resolved.companyName,
          },
          matchedEntityId: existingId,
          matchedEntityName: null,
          duplicateMatch: "nameCompany",
        });
        continue;
      }

      if (existingId && input.duplicateMode === "update") {
        if (importType === "nvq") {
          await updateAdminNvq(existingId, payload);
        } else {
          const key = REGISTER_KEY_BY_TYPE[importType];
          const { record: updated } = await updateAdminRegister(
            key,
            existingId,
            payload,
          );
          await triggerMatrixSyncAfterRegister(key, updated);
        }
        results.push({
          rowNumber: row.rowNumber,
          status: "Imported",
          messages: ["Updated existing row."],
          fields: {
            ...fields,
            candidateName: resolved.candidateName,
            company: resolved.companyName,
          },
          matchedEntityId: existingId,
          matchedEntityName: null,
          duplicateMatch: "nameCompany",
        });
        continue;
      }

      if (importType === "nvq") {
        await createAdminNvq(payload);
      } else {
        const key = REGISTER_KEY_BY_TYPE[importType];
        const { record: created } = await createAdminRegister(key, payload);
        await triggerMatrixSyncAfterRegister(key, created);
      }

      results.push({
        rowNumber: row.rowNumber,
        status: "Imported",
        messages: existingId
          ? ["Created additional row (duplicate mode = create)."]
          : ["Imported."],
        fields: {
          ...fields,
          candidateName: resolved.candidateName,
          company: resolved.companyName,
        },
        matchedEntityId: existingId,
        matchedEntityName: null,
        duplicateMatch: existingId ? "nameCompany" : null,
      });
    } catch (error) {
      results.push({
        rowNumber: row.rowNumber,
        status: "Error",
        messages: [
          error instanceof Error
            ? error.message
            : "Import failed for this row.",
        ],
        fields,
        matchedEntityId: null,
        matchedEntityName: null,
        duplicateMatch: null,
      });
    }
  }

  return results;
}

export { summarizeBulkRows, isRegisterImportType };
