/**
 * Portal navigation — same destinations as Next.js admin / customer sidebars.
 * SPFx uses in-web-part view switching (button click) instead of App Router.
 */

export type AdminViewId =
  | "dashboard"
  | "companies"
  | "workforce"
  | "training-matrix"
  | "training-records"
  | "nvq"
  | "documents"
  | "events"
  | "offers"
  | "permissions"
  | "automation"
  | "logs";

export type CustomerViewId =
  | "dashboard"
  | "training-matrix"
  | "candidates"
  | "training-records"
  | "nvq-progress"
  | "documents"
  | "events"
  | "offers"
  | "support";

export const ADMIN_NAV: ReadonlyArray<{ id: AdminViewId; label: string }> = [
  { id: "dashboard", label: "Dashboard" },
  { id: "companies", label: "Companies" },
  { id: "workforce", label: "Workforce" },
  { id: "training-matrix", label: "Training Matrix" },
  { id: "training-records", label: "Training Records" },
  { id: "nvq", label: "NVQ" },
  { id: "documents", label: "Documents" },
  { id: "events", label: "Events" },
  { id: "offers", label: "Offers" },
  { id: "permissions", label: "Permissions" },
  { id: "automation", label: "Automation" },
  { id: "logs", label: "Logs" },
];

export const CUSTOMER_NAV: ReadonlyArray<{
  id: CustomerViewId;
  label: string;
}> = [
  { id: "dashboard", label: "Dashboard" },
  { id: "training-matrix", label: "Training Matrix" },
  { id: "candidates", label: "Candidates" },
  { id: "training-records", label: "Training Records" },
  { id: "nvq-progress", label: "NVQ Progress" },
  { id: "documents", label: "Documents" },
  { id: "events", label: "Events" },
  { id: "offers", label: "Offers" },
  { id: "support", label: "Support" },
];
