import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import { toolError } from '../graph/client';
import { getAll } from '../graph/pager';

const MSG_SELECT =
  'id,subject,from,toRecipients,ccRecipients,receivedDateTime,' +
  'isRead,hasAttachments,bodyPreview,conversationId,importance';

const ACCOUNT_PARAM = z
  .string()
  .optional()
  .describe(
    'Account to act as (email address or display-name substring). ' +
      'Omit to use the default account. Use accounts_list to see available accounts.',
  );

export function registerMailTools(
  server: McpServer,
  getGraph: (account?: string) => Client,
): void {
  // ── 1. List messages ────────────────────────────────────────────────────────
  server.tool(
    'mail_list_messages',
    'List email messages from a mail folder, with optional full-text search or OData filter.',
    {
      folder: z
        .string()
        .default('inbox')
        .describe('Folder name or well-known name: inbox, sentitems, drafts, deleteditems, junkemail, archive'),
      search: z.string().optional().describe('Free-text search across subject, body, and sender'),
      filter: z.string().optional().describe('OData $filter expression, e.g. "isRead eq false"'),
      top: z.number().int().min(1).max(100).default(25).describe('Max messages to return (default 25)'),
      account: ACCOUNT_PARAM,
    },
    async ({ folder, search, filter, top, account }) => {
      const graph = getGraph(account);
      try {
        let req = graph
          .api(`/me/mailFolders/${folder}/messages`)
          .select(MSG_SELECT)
          .top(top)
          .orderby('receivedDateTime desc');
        if (search) req = req.query({ $search: `"${search}"` });
        if (filter) req = req.filter(filter);
        const result = await req.get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(result.value, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 2. Get message ──────────────────────────────────────────────────────────
  server.tool(
    'mail_get_message',
    'Retrieve the full content of a single email message, including body and attachment metadata.',
    {
      message_id: z.string().describe('The message ID (from mail_list_messages)'),
      account: ACCOUNT_PARAM,
    },
    async ({ message_id, account }) => {
      const graph = getGraph(account);
      try {
        const message = await graph
          .api(`/me/messages/${message_id}`)
          .select(`${MSG_SELECT},body,attachments,sentDateTime,replyTo,internetMessageId`)
          .expand('attachments($select=id,name,size,contentType,isInline)')
          .get();
        return { content: [{ type: 'text' as const, text: JSON.stringify(message, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 3. Send message ─────────────────────────────────────────────────────────
  server.tool(
    'mail_send',
    'Send a new email message from the specified account.',
    {
      to: z.array(z.string()).min(1).describe('Recipient email addresses'),
      subject: z.string().describe('Email subject line'),
      body: z.string().describe('Email body (plain text)'),
      cc: z.array(z.string()).optional().describe('CC email addresses'),
      save_to_sent: z.boolean().default(true).describe('Save a copy to Sent Items'),
      account: ACCOUNT_PARAM,
    },
    async ({ to, subject, body, cc, save_to_sent, account }) => {
      const graph = getGraph(account);
      try {
        const message: Record<string, unknown> = {
          subject,
          body: { contentType: 'text', content: body },
          toRecipients: to.map((address) => ({ emailAddress: { address } })),
        };
        if (cc?.length) {
          message.ccRecipients = cc.map((address) => ({ emailAddress: { address } }));
        }
        await graph.api('/me/sendMail').post({ message, saveToSentItems: save_to_sent });
        return { content: [{ type: 'text' as const, text: 'Message sent.' }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 4. Reply ────────────────────────────────────────────────────────────────
  server.tool(
    'mail_reply',
    'Reply or reply-all to an existing email message.',
    {
      message_id: z.string().describe('The message ID to reply to'),
      body: z.string().describe('Your reply body (plain text)'),
      reply_all: z.boolean().default(false).describe('true = reply to all; false = reply to sender only'),
      account: ACCOUNT_PARAM,
    },
    async ({ message_id, body, reply_all, account }) => {
      const graph = getGraph(account);
      try {
        const action = reply_all ? 'replyAll' : 'reply';
        await graph.api(`/me/messages/${message_id}/${action}`).post({
          message: { body: { contentType: 'text', content: body } },
        });
        return { content: [{ type: 'text' as const, text: 'Reply sent.' }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 5. Move message ─────────────────────────────────────────────────────────
  server.tool(
    'mail_move',
    'Move an email message to a different mail folder.',
    {
      message_id: z.string().describe('The message ID to move'),
      destination_folder: z
        .string()
        .describe('Destination folder name or well-known name (inbox, archive, deleteditems, junkemail)'),
      account: ACCOUNT_PARAM,
    },
    async ({ message_id, destination_folder, account }) => {
      const graph = getGraph(account);
      try {
        await graph.api(`/me/messages/${message_id}/move`).post({ destinationId: destination_folder });
        return { content: [{ type: 'text' as const, text: `Message moved to "${destination_folder}".` }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );

  // ── 6. List folders ─────────────────────────────────────────────────────────
  server.tool(
    'mail_list_folders',
    'List mail folders for an account, including unread and total item counts.',
    { account: ACCOUNT_PARAM },
    async ({ account }) => {
      const graph = getGraph(account);
      try {
        const folders = await getAll(graph, '/me/mailFolders', {
          $select: 'id,displayName,unreadItemCount,totalItemCount,isHidden',
        });
        return { content: [{ type: 'text' as const, text: JSON.stringify(folders, null, 2) }] };
      } catch (e) {
        return toolError(e);
      }
    },
  );
}
