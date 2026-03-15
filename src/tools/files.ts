import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { toolError } from '../graph/client';

const ITEM_SELECT = 'id,name,size,lastModifiedDateTime,file,folder,webUrl,parentReference';

const ACCOUNT_PARAM = z
  .string()
  .optional()
  .describe(
    'Account to act as (email address or display-name substring). ' +
      'Omit to use the default account. Use accounts_list to see available accounts.',
  );

export function registerFilesTools(
  server: McpServer,
  getGraph: (account?: string) => Client,
): void {
  // ── 1. List files/folders ───────────────────────────────────────────────────
  server.tool(
    'files_list',
    'List files and folders in OneDrive or a SharePoint site document library. Omit path to list the root.',
    {
      path: z
        .string()
        .optional()
        .describe('Folder path relative to root, e.g. "Documents/Reports". Empty = root.'),
      site: z
        .string()
        .optional()
        .describe(
          'SharePoint site address to access instead of personal OneDrive. ' +
          'Format: "hostname:/sites/SiteName" e.g. "medmincouk.sharepoint.com:/sites/MedminSoftwareDevelopment". ' +
          'Omit to use personal OneDrive.',
        ),
      top: z.number().int().min(1).max(200).default(50).describe('Max items to return (default 50)'),
      account: ACCOUNT_PARAM,
    },
    async ({ path, site, top, account }) => {
      const graph = getGraph(account);
      try {
        let endpoint: string;
        if (site) {
          // Resolve site to get its stable ID, then use that for the drive query
          const siteData = await graph.api(`/sites/${site}`).select('id').get() as { id: string };
          const siteId = siteData.id;
          endpoint = path
            ? `/sites/${siteId}/drive/root:/${path}:/children`
            : `/sites/${siteId}/drive/root/children`;
        } else {
          endpoint = path
            ? `/me/drive/root:/${encodeURIComponent(path)}:/children`
            : '/me/drive/root/children';
        }
        const result = await graph.api(endpoint).select(ITEM_SELECT).top(top).get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 2. Get file content ─────────────────────────────────────────────────────
  server.tool(
    'files_get_content',
    'Download a file from OneDrive or SharePoint. Text files are returned as plain text; binary files as "base64:<data>".',
    {
      path: z.string().optional().describe('File path relative to root, e.g. "Documents/notes.txt"'),
      item_id: z.string().optional().describe('OneDrive/SharePoint item ID — alternative to path'),
      site: z
        .string()
        .optional()
        .describe(
          'SharePoint site address (used with path). ' +
          'Format: "hostname:/sites/SiteName" e.g. "medmincouk.sharepoint.com:/sites/MedminSoftwareDevelopment". ' +
          'Omit to use personal OneDrive.',
        ),
      drive_id: z
        .string()
        .optional()
        .describe('Drive ID — use with item_id when the item is on a SharePoint drive, not personal OneDrive.'),
      account: ACCOUNT_PARAM,
    },
    async ({ path, item_id, site, drive_id, account }) => {
      const graph = getGraph(account);
      try {
        if (!path && !item_id) {
          return {
            isError: true as const,
            content: [{ type: 'text' as const, text: 'Error: supply either path or item_id.' }],
          };
        }
        let endpoint: string;
        if (item_id && drive_id) {
          endpoint = `/drives/${drive_id}/items/${item_id}/content`;
        } else if (item_id) {
          endpoint = `/me/drive/items/${item_id}/content`;
        } else if (site) {
          const siteData = await graph.api(`/sites/${site}`).select('id').get() as { id: string };
          endpoint = `/sites/${siteData.id}/drive/root:/${path}:/content`;
        } else {
          endpoint = `/me/drive/root:/${encodeURIComponent(path!)}:/content`;
        }
        const raw = await graph.api(endpoint).responseType(ResponseType.ARRAYBUFFER).get();
        const buffer = Buffer.from(raw as ArrayBuffer);
        const text = buffer.toString('utf-8');
        const isBinary = text.includes('\uFFFD');
        return {
          content: [{ type: 'text' as const, text: isBinary ? `base64:${buffer.toString('base64')}` : text }],
        };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 3. Upload / overwrite file ──────────────────────────────────────────────
  server.tool(
    'files_upload',
    'Upload or overwrite a file in OneDrive (max ~4 MB). Prefix binary content with "base64:".',
    {
      path: z.string().describe('Destination path in OneDrive, e.g. "Documents/notes.txt"'),
      content: z.string().describe('File content as plain text, or binary as "base64:<data>".'),
      account: ACCOUNT_PARAM,
    },
    async ({ path, content, account }) => {
      const graph = getGraph(account);
      try {
        const isBase64 = content.startsWith('base64:');
        const buffer = isBase64
          ? Buffer.from(content.slice(7), 'base64')
          : Buffer.from(content, 'utf-8');
        const item = await graph.api(`/me/drive/root:/${encodeURIComponent(path)}:/content`).put(buffer);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                { id: item.id, name: item.name, size: item.size, webUrl: item.webUrl },
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

  // ── 4. Search ───────────────────────────────────────────────────────────────
  server.tool(
    'files_search',
    "Full-text search across all files in an account's OneDrive.",
    {
      query: z.string().describe('Search query string'),
      top: z.number().int().min(1).max(50).default(20).describe('Max results (default 20)'),
      account: ACCOUNT_PARAM,
    },
    async ({ query, top, account }) => {
      const graph = getGraph(account);
      try {
        const result = await graph
          .api(`/me/drive/root/search(q='${encodeURIComponent(query)}')`)
          .select(ITEM_SELECT)
          .top(top)
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 5. Create sharing link ──────────────────────────────────────────────────
  server.tool(
    'files_get_sharing_link',
    'Create a shareable link for a OneDrive file or folder.',
    {
      path: z.string().optional().describe('File path relative to OneDrive root'),
      item_id: z.string().optional().describe('OneDrive item ID — alternative to path'),
      type: z.enum(['view', 'edit']).default('view').describe('"view" = read-only; "edit" = read-write'),
      scope: z
        .enum(['anonymous', 'organization'])
        .default('organization')
        .describe('"anonymous" = anyone with link; "organization" = org members only'),
      account: ACCOUNT_PARAM,
    },
    async ({ path, item_id, type, scope, account }) => {
      const graph = getGraph(account);
      try {
        if (!path && !item_id) {
          return {
            isError: true as const,
            content: [{ type: 'text' as const, text: 'Error: supply either path or item_id.' }],
          };
        }
        let id = item_id;
        if (!id) {
          const item = await graph.api(`/me/drive/root:/${encodeURIComponent(path!)}`).select('id').get();
          id = item.id as string;
        }
        const link = await graph.api(`/me/drive/items/${id}/createLink`).post({ type, scope });
        return { content: [{ type: 'text' as const, text: JSON.stringify(link, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
