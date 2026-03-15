import { homedir } from 'node:os';
import { join } from 'node:path';

export interface Config {
  clientId: string;
  tenantId: string;
  /** Absolute path to the persistent token cache file */
  cachePath: string;
}

export function loadConfig(): Config {
  const clientId = process.env.AZURE_CLIENT_ID?.trim();
  const tenantId = process.env.AZURE_TENANT_ID?.trim() || 'common';
  const cachePath =
    process.env.MCP_CACHE_PATH?.trim() ||
    join(homedir(), '.mcp-o365-token-cache.json');

  if (!clientId) {
    process.stderr.write(
      [
        'Error: AZURE_CLIENT_ID is not set.',
        'Create an Azure app registration (public client) and set:',
        '  AZURE_CLIENT_ID=<your-client-id>',
        '  AZURE_TENANT_ID=common   (or your tenant GUID)',
        'See README.md for full setup instructions.',
        '',
      ].join('\n'),
    );
    process.exit(1);
  }

  return { clientId, tenantId, cachePath };
}
