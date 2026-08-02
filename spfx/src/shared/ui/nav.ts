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
  | "bulk-upload"
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
  { id: "dashboard", label: "Home" },
  { id: "companies", label: "Companies" },
  { id: "workforce", label: "Workforce" },
  { id: "training-matrix", label: "Matrix" },
  { id: "training-records", label: "Registers" },
  { id: "nvq", label: "NVQ" },
  { id: "documents", label: "Documents" },
  { id: "events", label: "Calendar" },
  { id: "offers", label: "Offers" },
  { id: "permissions", label: "Permissions" },
  { id: "bulk-upload", label: "Bulk upload" },
  { id: "automation", label: "Automation" },
  { id: "logs", label: "Audit Log" },
];

export const CUSTOMER_NAV: ReadonlyArray<{
  id: CustomerViewId;
  label: string;
}> = [
  { id: "training-matrix", label: "Training Matrix" },
  { id: "dashboard", label: "Dashboard" },
  { id: "candidates", label: "Candidates / Workforce" },
  { id: "training-records", label: "Training Records" },
  { id: "nvq-progress", label: "NVQ Progress" },
  { id: "documents", label: "Documents" },
  { id: "events", label: "Events / Bookings" },
  { id: "offers", label: "Offers" },
  { id: "support", label: "Support" },
];
