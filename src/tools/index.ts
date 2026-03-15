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

export function registerAllTools(
  server: McpServer,
  accountManager: AccountManager,
  getGraph: (account?: string) => Client,
): void {
  registerAccountsTools(server, accountManager);
  registerCalendarTools(server, getGraph);
  registerMailTools(server, getGraph);
  registerFilesTools(server, getGraph);
  registerContactsTools(server, getGraph);
  registerUserTools(server, getGraph);
  registerMeetingsTools(server, getGraph);
}
