import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { toolError } from '../graph/client';
import { getAll } from '../graph/pager';

const CONTACT_SELECT =
  'id,displayName,emailAddresses,businessPhones,mobilePhone,jobTitle,companyName,department';

const ACCOUNT_PARAM = z
  .string()
  .optional()
  .describe(
    'Account to act as (email address or display-name substring). ' +
      'Omit to use the default account. Use accounts_list to see available accounts.',
  );

export function registerContactsTools(
  server: McpServer,
  getGraph: (account?: string) => Client,
): void {
  // ── 1. List contacts ────────────────────────────────────────────────────────
  server.tool(
    'contacts_list',
    'List personal Outlook contacts for an account, sorted alphabetically.',
    {
      top: z.number().int().min(1).max(100).default(50).describe('Max contacts to return (default 50)'),
      account: ACCOUNT_PARAM,
    },
    async ({ top, account }) => {
      const graph = getGraph(account);
      try {
        const result = await graph
          .api('/me/contacts')
          .select(CONTACT_SELECT)
          .top(top)
          .orderby('displayName')
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 2. Search personal contacts ─────────────────────────────────────────────
  server.tool(
    'contacts_search',
    'Search personal Outlook contacts by display name or email address prefix.',
    {
      query: z.string().describe('Name or email prefix to search for'),
      account: ACCOUNT_PARAM,
    },
    async ({ query, account }) => {
      const graph = getGraph(account);
      try {
        const safe = query.replace(/'/g, "''");
        const result = await graph
          .api('/me/contacts')
          .select(CONTACT_SELECT)
          .filter(
            `startswith(displayName,'${safe}') or emailAddresses/any(e:startswith(e/address,'${safe}'))`,
          )
          .top(25)
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 3. People / org directory search ────────────────────────────────────────
  server.tool(
    'people_search',
    'Search the organisation directory using the People API. Results are relevance-ranked based on communication history, org chart, and name match.',
    {
      query: z.string().describe('Name, alias, or keyword to search for'),
      top: z.number().int().min(1).max(50).default(10).describe('Max results (default 10)'),
      account: ACCOUNT_PARAM,
    },
    async ({ query, top, account }) => {
      const graph = getGraph(account);
      try {
        const result = await graph
          .api('/me/people')
          .query({ $search: `"${query}"` })
          .select(
            'id,displayName,scoredEmailAddresses,jobTitle,department,officeLocation,phones,userPrincipalName',
          )
          .top(top)
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
