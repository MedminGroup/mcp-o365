import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { toolError } from '../graph/client';

const ACCOUNT_PARAM = z
  .string()
  .optional()
  .describe(
    'Account to act as (email address or display-name substring). ' +
      'Omit to use the default account. Use accounts_list to see available accounts.',
  );

export function registerUserTools(
  server: McpServer,
  getGraph: (account?: string) => Client,
): void {
  server.tool(
    'user_get_profile',
    "Get a signed-in account's profile: name, email, job title, department, office location, and phone numbers.",
    { account: ACCOUNT_PARAM },
    async ({ account }) => {
      const graph = getGraph(account);
      try {
        const user = await graph
          .api('/me')
          .select(
            'id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones',
          )
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(user, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  server.tool(
    'user_get_photo_url',
    'Get profile photo metadata for the signed-in account, or another user by email.',
    {
      email: z.string().optional().describe("Email/UPN of another user; omit for the account's own photo"),
      account: ACCOUNT_PARAM,
    },
    async ({ email, account }) => {
      const graph = getGraph(account);
      try {
        const base = email ? `/users/${encodeURIComponent(email)}` : '/me';
        const meta = await graph.api(`${base}/photo`).get();
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                { ...meta, binaryEndpoint: `${base}/photo/$value` },
                null,
                2,
              ),
            },
          ],
        };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
