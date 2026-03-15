import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import type { AccountManager } from '../auth';
import { toolError } from '../graph/client';

export function registerAccountsTools(server: McpServer, accountManager: AccountManager): void {
  // ── 1. List signed-in accounts ──────────────────────────────────────────────
  server.tool(
    'accounts_list',
    'List all Microsoft 365 accounts currently signed in. Use the username value as the "account" parameter in other tools to target a specific account.',
    {},
    async () => {
      try {
        const accounts = await accountManager.listAccounts();
        if (accounts.length === 0) {
          return {
            content: [
              {
                type: 'text' as const,
                text: 'No accounts signed in yet. Use accounts_add to sign in.',
              },
            ],
          };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(accounts, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 2. Start sign-in (returns URL + code immediately) ───────────────────────
  server.tool(
    'accounts_add',
    [
      'Start adding a new Microsoft 365 account via device-code sign-in.',
      'Returns a URL and short code — open the URL in your browser and enter the code to sign in.',
      'After signing in, call accounts_complete to finish and confirm the account was added.',
      'You can add multiple accounts from different tenants; all share a single app registration.',
    ].join(' '),
    {},
    async () => {
      try {
        const info = await accountManager.startDeviceCodeFlow();
        return {
          content: [
            {
              type: 'text' as const,
              text: [
                '**Microsoft 365 sign-in — action required**',
                '',
                `1. Open this URL in your browser:  ${info.verificationUri}`,
                `2. Enter this code when prompted:   **${info.userCode}**`,
                '',
                `Code expires in ${info.expiresInMinutes} minutes.`,
                '',
                'Once you have signed in, call **accounts_complete** to confirm.',
              ].join('\n'),
            },
          ],
        };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 3. Complete sign-in (awaits background polling) ─────────────────────────
  server.tool(
    'accounts_complete',
    'Complete a sign-in started by accounts_add. Call this after you have visited the URL and entered the code in your browser. Returns the account that was just added.',
    {},
    async () => {
      try {
        const account = await accountManager.completeDeviceCodeFlow();
        return {
          content: [
            {
              type: 'text' as const,
              text: [
                '**Sign-in complete!**',
                '',
                JSON.stringify(account, null, 2),
                '',
                'Use accounts_list to see all signed-in accounts.',
              ].join('\n'),
            },
          ],
        };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 4. Remove a signed-in account ───────────────────────────────────────────
  server.tool(
    'accounts_remove',
    'Remove a signed-in account from the local token cache. The account will need to sign in again via accounts_add if used in the future.',
    {
      account: z
        .string()
        .describe('Email address or display-name substring of the account to remove'),
    },
    async ({ account }) => {
      try {
        const removed = await accountManager.removeAccount(account);
        return {
          content: [
            {
              type: 'text' as const,
              text: `Removed: ${removed.username} (${removed.name})`,
            },
          ],
        };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
