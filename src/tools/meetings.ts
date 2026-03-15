import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { toolError } from '../graph/client';

const ACCOUNT_PARAM = z
  .string()
  .optional()
  .describe(
    'Account to act as (email address or display-name substring). ' +
      'Omit to use the default account. Use accounts_list to see available accounts.',
  );

export function registerMeetingsTools(
  server: McpServer,
  getGraph: (account?: string) => Client,
): void {
  // ── 1. Get online meeting by join URL ────────────────────────────────────────
  server.tool(
    'meetings_get_by_join_url',
    'Look up a Teams online meeting by its join URL (the joinWebUrl field from a calendar event). Returns the meeting ID required for transcript calls.',
    {
      join_url: z.string().describe('The Teams join URL from the calendar event joinWebUrl field'),
      account: ACCOUNT_PARAM,
    },
    async ({ join_url, account }) => {
      const graph = getGraph(account);
      try {
        const result = await graph
          .api('/me/onlineMeetings')
          .filter(`joinWebUrl eq '${join_url.replace(/'/g, "''")}'`)
          .get();
        const meetings = (result.value ?? []) as unknown[];
        if (meetings.length === 0) {
          return {
            isError: true as const,
            content: [{ type: 'text' as const, text: 'No online meeting found for that join URL. The meeting may not have been organised by this account.' }],
          };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(meetings[0], null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 2. List transcripts for a meeting ────────────────────────────────────────
  server.tool(
    'meetings_list_transcripts',
    'List all available transcripts for a Teams online meeting, ordered by creation time. For recurring meetings, filter by createdDateTime to find a specific occurrence.',
    {
      meeting_id: z.string().describe('The online meeting ID (from meetings_get_by_join_url)'),
      account: ACCOUNT_PARAM,
    },
    async ({ meeting_id, account }) => {
      const graph = getGraph(account);
      try {
        const result = await graph
          .api(`/me/onlineMeetings/${meeting_id}/transcripts`)
          .get();
        const transcripts = result.value ?? result;
        if (Array.isArray(transcripts) && transcripts.length === 0) {
          return {
            content: [{ type: 'text' as const, text: 'No transcripts found for this meeting. Transcription must be started manually during the meeting in Teams.' }],
          };
        }
        return { content: [{ type: 'text' as const, text: JSON.stringify(transcripts, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 3. Download transcript content ───────────────────────────────────────────
  server.tool(
    'meetings_get_transcript',
    'Download the VTT transcript content for a specific transcript. Returns raw VTT text with speaker labels and timestamps, ready for analysis.',
    {
      meeting_id: z.string().describe('The online meeting ID'),
      transcript_id: z.string().describe('The transcript ID (from meetings_list_transcripts)'),
      account: ACCOUNT_PARAM,
    },
    async ({ meeting_id, transcript_id, account }) => {
      const graph = getGraph(account);
      try {
        const raw = await graph
          .api(`/me/onlineMeetings/${meeting_id}/transcripts/${transcript_id}/content`)
          .header('Accept', 'text/vtt')
          .responseType(ResponseType.TEXT)
          .get();
        return { content: [{ type: 'text' as const, text: String(raw) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
