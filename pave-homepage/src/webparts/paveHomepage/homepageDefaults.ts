/** Shared defaults for property pane + first drop on a page. */
export function buildHomepageDefaults(webAbsoluteUrl: string): {
  heroHeadline: string;
  heroSubtitle: string;
  greetingNameOverride: string;
  nporsRegisterUrl: string;
  workforceUrl: string;
  eusrUrl: string;
  customerDocsUrl: string;
  adminDocsUrl: string;
  documentsUrl: string;
  eventsAddUrl: string;
  heroImageUrl: string;
  workforceImageUrl: string;
  eusrImageUrl: string;
  customerDocsImageUrl: string;
  adminDocsImageUrl: string;
} {
  const base = webAbsoluteUrl.replace(/\/$/, "");
  const lists = `${base}/Lists`;
  const siteAssets = `${base}/SiteAssets`;

  return {
    heroHeadline: "NPORS renewals are due this month",
    heroSubtitle:
      "Check the register to see which operators need re-certification",
    greetingNameOverride: "",
    nporsRegisterUrl: `${lists}/NPORS%20Register/AllItems.aspx`,
    workforceUrl: `${lists}/Workforce%20List/AllItems.aspx`,
    eusrUrl: `${lists}/EUSR%20Register/AllItems.aspx`,
    customerDocsUrl: `${base}/CustomerDocuments/Forms/AllItems.aspx`,
    adminDocsUrl: `${base}/Shared%20Documents/Forms/AllItems.aspx`,
    documentsUrl: `${base}/Shared%20Documents/Forms/AllItems.aspx`,
    eventsAddUrl: `${lists}/Events/NewForm.aspx`,
    heroImageUrl: `${siteAssets}/pave-hero-placeholder.jpg`,
    workforceImageUrl: "",
    eusrImageUrl: "",
    customerDocsImageUrl: "",
    adminDocsImageUrl: "",
  };
}
