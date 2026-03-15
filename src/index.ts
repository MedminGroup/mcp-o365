import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { loadConfig } from './config';
import { AccountManager } from './auth';
import { createGraphClient } from './graph/client';
import { registerAllTools } from './tools';
import { runSetup } from './setup';

const version = '1.0.12'; // keep in sync with package.json

async function main(): Promise<void> {
  const config = loadConfig();

  if (process.argv.includes('--setup')) {
    await runSetup(config.clientId, config.tenantId, config.cachePath);
    process.exit(0);
  }

  // Single MSAL instance, shared token cache — supports multiple signed-in accounts
  const accountManager = new AccountManager(config);

  // Per-call Graph client factory — each call gets a client scoped to the right account
  const getGraph = (account?: string) => createGraphClient(accountManager, account);

  const server = new McpServer({ name: 'mcp-o365', version });

  registerAllTools(server, accountManager, getGraph);

  const transport = new StdioServerTransport();
  await server.connect(transport);

  process.stderr.write(`mcp-o365 v${version} ready (stdio)\n`);
}

main().catch((err: unknown) => {
  const msg = err instanceof Error ? err.message : String(err);
  process.stderr.write(`Fatal: ${msg}\n`);
  process.exit(1);
});
