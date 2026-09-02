import "server-only";

import { revalidatePath } from "next/cache";

import { revalidateSharePointLists } from "@/lib/cache/sharePointCache";
import type { SharePointListKey } from "@/lib/schema/sharepointSchema";

/**
 * Lists that feed Workforce profile, Candidate profile, and Training summary.
 * A write to any register / matrix / NVQ can change all three surfaces.
 */
const TRAINING_SURFACE_LISTS: SharePointListKey[] = [
  "nporsRegister",
  "eusrRegister",
  "nrswaRegister",
  "inHouseCertificates",
  "nvqRegister",
  "trainingMatrixExample",
  "trainingMatrix",
  "trainingMatrixCategoryRecords",
  "workforce",
];

const TRAINING_SURFACE_PATHS: Array<{ path: string; type?: "layout" | "page" }> =
  [
    { path: "/admin/workforce", type: "layout" },
    { path: "/admin/training-matrix" },
    { path: "/admin/training-records", type: "layout" },
    { path: "/admin/nvq" },
    { path: "/customer", type: "layout" },
  ];

/**
 * After Training Matrix / NPORS / EUSR / Streetworks / In-House / NVQ changes,
 * drop SharePoint + App Router caches so profile and summary pages do not
 * require a manual refresh.
 */
export function revalidateTrainingSurfaces(): void {
  revalidateSharePointLists(TRAINING_SURFACE_LISTS);
  for (const entry of TRAINING_SURFACE_PATHS) {
    revalidatePath(entry.path, entry.type);
  }
}
