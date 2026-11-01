/**
 * Event documents live in the existing Customer Documents library, in the
 * event's company folder:
 *   {Company folder}/Company Documents/Events/{event id}/{file}
 * The folder is the link to the event (its id never changes, unlike its
 * title), so no SharePoint schema change is needed. Pure helpers — tested in
 * scripts/test-event-documents.mjs.
 */

/** Folders under "Company Documents" for one event. */
export function eventDocumentSubfolders(eventId: string): string[] {
  return ["Events", String(eventId).trim()];
}

/**
 * Event id a document belongs to, from its SharePoint folder path
 * (`FileDirRef`), or null when it is not an event document.
 */
export function eventIdFromDocumentFolder(
  fileDirRef: string | null | undefined,
): string | null {
  const match = (fileDirRef ?? "")
    .trim()
    .replace(/\/+$/, "")
    .match(/\/Company Documents\/Events\/(\d+)$/i);
  return match?.[1] ?? null;
}
