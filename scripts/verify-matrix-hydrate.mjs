/**
 * Verify Training Matrix Update rows hydrate with candidate names + CSCS dates.
 *   node --env-file=.env.local scripts/verify-matrix-hydrate.mjs
 */
async function main() {
  // Use Graph directly mirroring the app hydrate fix.
  const body = new URLSearchParams({
    client_id: process.env.AZURE_CLIENT_ID,
    client_secret: process.env.AZURE_CLIENT_SECRET,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const tok = (
    await (
      await fetch(
        `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`,
        {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body,
        },
      )
    ).json()
  ).access_token;
  const site = process.env.SHAREPOINT_SITE_ID.replace(/\/+$/, "");
  const root =
    "/sites/" +
    (site.includes(":/") ? (site.endsWith(":") ? site : site + ":") : site);
  const listId = process.env.SHAREPOINT_TRAINING_MATRIX_EXAMPLE_LIST_ID;

  const expand = await (
    await fetch(
      `https://graph.microsoft.com/v1.0${root}/lists/${listId}/items?$expand=fields&$top=5`,
      { headers: { Authorization: `Bearer ${tok}` } },
    )
  ).json();

  const emptyExpand = (expand.value || []).every(
    (i) => !i.fields || Object.keys(i.fields).filter((k) => !k.startsWith("@")).length === 0,
  );
  console.log("collection_expand_empty", emptyExpand, "n", expand.value?.length);

  const ids = (expand.value || []).map((i) => i.id).slice(0, 5);
  let named = 0;
  for (const id of ids) {
    const f = await (
      await fetch(
        `https://graph.microsoft.com/v1.0${root}/lists/${listId}/items/${id}/fields`,
        { headers: { Authorization: `Bearer ${tok}` } },
      )
    ).json();
    if (f.Title) named += 1;
    console.log(id, f.Title, f.CSCSExpiry?.slice?.(0, 10) || f.CSCSExpiry);
  }
  console.log("hydrated_named", named, "/", ids.length);
  if (named === 0) process.exit(1);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
