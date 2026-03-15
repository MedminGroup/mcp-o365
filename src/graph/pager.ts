import type { Client } from '@microsoft/microsoft-graph-client';

/**
 * Transparently pages through all results from a Graph API collection endpoint.
 *
 * Graph returns at most 999 items per page (often 100 by default) and includes
 * an `@odata.nextLink` URL when there are more pages. This utility follows
 * those links automatically so callers always get the full result set.
 *
 * @param client  Authenticated Graph client
 * @param path    Graph API path, e.g. '/me/calendars'
 * @param query   Optional OData query params applied to the first request only
 */
export async function getAll<T>(
  client: Client,
  path: string,
  query?: Record<string, string | number>,
): Promise<T[]> {
  const results: T[] = [];

  let req = client.api(path);
  if (query) {
    req = req.query(query);
  }

  let page: { value?: T[]; '@odata.nextLink'?: string } | null = await req.get();

  while (page) {
    if (Array.isArray(page.value)) {
      results.push(...page.value);
    }
    if (page['@odata.nextLink']) {
      page = await client.api(page['@odata.nextLink']).get();
    } else {
      break;
    }
  }

  return results;
}
