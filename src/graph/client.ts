import { Client, ResponseType } from '@microsoft/microsoft-graph-client';
import type { AccountManager } from '../auth';

export { ResponseType };

/**
 * Create a Graph client scoped to a specific account.
 * `accountHint` is matched against the cached account's username or display
 * name (case-insensitive substring). Omit to use the default (first) account.
 */
export function createGraphClient(
  accountManager: AccountManager,
  accountHint?: string,
): Client {
  return Client.init({
    authProvider: (done) => {
      accountManager
        .getTokenProvider(accountHint)()
        .then((token) => done(null, token))
        .catch((err: Error) => done(err, null));
    },
  });
}

// ─── Error helper ─────────────────────────────────────────────────────────────

type ToolErrorResult = {
  isError: true;
  content: [{ type: 'text'; text: string }];
};

/**
 * Converts any thrown value into the MCP tool-error shape.
 * Graph SDK errors carry a `.body` JSON string with the real message.
 */
export function toolError(error: unknown): ToolErrorResult {
  let message: string;

  if (error instanceof Error) {
    const raw = (error as Error & { body?: string }).body;
    if (raw) {
      try {
        const parsed = JSON.parse(raw) as { error?: { message?: string } };
        message = parsed.error?.message ?? raw;
      } catch {
        message = raw;
      }
    } else {
      message = error.message;
    }
  } else {
    message = String(error);
  }

  return {
    isError: true,
    content: [{ type: 'text', text: `Graph API error: ${message}` }],
  };
}
