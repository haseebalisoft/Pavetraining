import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx as SpSpfx } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as GraphSpfx } from "@pnp/graph";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/graph/users";
import "@pnp/graph/calendars";

let _sp: SPFI | undefined;
let _graph: GraphFI | undefined;
let _contextKey: string | undefined;

function contextKey(context: WebPartContext): string {
  return [
    context.pageContext.web.absoluteUrl,
    context.pageContext.user?.loginName || "",
  ].join("|");
}

/**
 * Returns a cached PnPjs SharePoint client bound to the current SPFx context.
 * Call once from the web part `onInit` (or before first service use).
 */
export function getSP(context: WebPartContext): SPFI {
  const key = contextKey(context);
  if (!_sp || _contextKey !== key) {
    _sp = spfi().using(SpSpfx(context));
    _contextKey = key;
  }
  return _sp;
}

/**
 * Returns a cached PnPjs Graph client bound to the current SPFx context.
 */
export function getGraph(context: WebPartContext): GraphFI {
  const key = contextKey(context);
  if (!_graph || _contextKey !== key) {
    _graph = graphfi().using(GraphSpfx(context));
    _contextKey = key;
  }
  return _graph;
}

/** Clears cached clients (useful in tests / after dispose). */
export function resetSPContext(): void {
  _sp = undefined;
  _graph = undefined;
  _contextKey = undefined;
}
