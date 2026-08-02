import { WebPartContext } from "@microsoft/sp-webpart-base";

import type { CalendarEvent } from "../models";
import { FIELDS, LIST_TITLES } from "./listSchema";
import { getGraph, getSP } from "./SPContext";

function asString(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value.trim();
  return String(value);
}

function asDateString(value: unknown): string | null {
  const raw = asString(value);
  if (!raw) return null;
  const date = new Date(raw);
  return Number.isNaN(date.getTime()) ? raw : date.toISOString();
}

function mapSharePointEvent(item: Record<string, unknown>): CalendarEvent {
  const f = FIELDS.events;
  return {
    id: asString(item[f.id] ?? item.Id ?? item.ID),
    title: asString(item[f.title]),
    start: asDateString(item[f.start]),
    end: asDateString(item[f.end]),
    location: asString(item[f.location]),
  };
}

async function getSharePointEvents(
  context: WebPartContext,
  top: number
): Promise<CalendarEvent[]> {
  const sp = getSP(context);
  const f = FIELDS.events;
  const now = new Date().toISOString();

  try {
    const items = await sp.web.lists
      .getByTitle(LIST_TITLES.events)
      .items.select(f.id, f.title, f.start, f.end, f.location)
      .filter(`${f.start} ge datetime'${now}'`)
      .orderBy(f.start, true)
      .top(Math.max(1, top))();

    return (items as Record<string, unknown>[]).map(mapSharePointEvent);
  } catch {
    // List exists but filter/order may fail on empty date fields — fall back unfiltered.
    const items = await sp.web.lists
      .getByTitle(LIST_TITLES.events)
      .items.select(f.id, f.title, f.start, f.end, f.location)
      .orderBy(f.start, true)
      .top(Math.max(1, top))();

    return (items as Record<string, unknown>[]).map(mapSharePointEvent);
  }
}

async function getOutlookEvents(
  context: WebPartContext,
  top: number
): Promise<CalendarEvent[]> {
  const graph = getGraph(context);
  const start = new Date();
  const end = new Date();
  end.setDate(end.getDate() + 60);

  const events = await graph.me.calendarView(
    start.toISOString(),
    end.toISOString()
  )
    .select("id", "subject", "start", "end", "location")
    .orderBy("start/dateTime", true)
    .top(Math.max(1, top))();

  return (events as Array<Record<string, unknown>>).map((item) => {
    const startObj = item.start as { dateTime?: string } | undefined;
    const endObj = item.end as { dateTime?: string } | undefined;
    const locationObj = item.location as { displayName?: string } | undefined;
    return {
      id: asString(item.id),
      title: asString(item.subject),
      start: asDateString(startObj && startObj.dateTime),
      end: asDateString(endObj && endObj.dateTime),
      location: asString(locationObj && locationObj.displayName),
    };
  });
}

/**
 * Upcoming events from SharePoint "Events" list.
 * Falls back to the signed-in user's Outlook calendar via Graph.
 */
export class EventsService {
  public static async getUpcoming(
    context: WebPartContext,
    top: number = 6
  ): Promise<CalendarEvent[]> {
    try {
      return await getSharePointEvents(context, top);
    } catch (spError) {
      console.warn(
        '[EventsService.getUpcoming] SharePoint "Events" unavailable — trying Outlook calendar',
        spError
      );
      try {
        return await getOutlookEvents(context, top);
      } catch (graphError) {
        console.error(
          "[EventsService.getUpcoming] Outlook calendar fallback failed",
          graphError
        );
        throw graphError;
      }
    }
  }
}
