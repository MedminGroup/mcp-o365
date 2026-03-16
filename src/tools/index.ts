import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import type { AccountManager } from '../auth';
import { registerAccountsTools } from './accounts';
import { registerCalendarTools } from './calendar';
import { registerMailTools } from './mail';
import { registerFilesTools } from './files';
import { registerContactsTools } from './contacts';
import { registerUserTools } from './user';
import { registerMeetingsTools } from './meetings';
import { registerGuideTools } from './guide';

type RegisterFn = () => void;

function safeRegister(name: string, fn: RegisterFn): void {
  try {
    fn();
    process.stderr.write(`[mcp-o365] registered: ${name}\n`);
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    process.stderr.write(`[mcp-o365] FAILED to register ${name}: ${msg}\n`);
  }
}

export function registerAllTools(
  server: McpServer,
  accountManager: AccountManager,
  getGraph: (account?: string) => Client,
): void {
  safeRegister('accounts',  () => registerAccountsTools(server, accountManager));
  safeRegister('calendar',  () => registerCalendarTools(server, getGraph));
  safeRegister('mail',      () => registerMailTools(server, getGraph));
  safeRegister('files',     () => registerFilesTools(server, getGraph));
  safeRegister('contacts',  () => registerContactsTools(server, getGraph));
  safeRegister('user',      () => registerUserTools(server, getGraph));
  safeRegister('meetings',  () => registerMeetingsTools(server, getGraph));
  safeRegister('guide',     () => registerGuideTools(server));
}
