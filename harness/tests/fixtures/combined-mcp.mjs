#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';

const server = new Server({ name: 'synthetic-moff-fixture', version: '1' }, {
  capabilities: { tools: {} },
  instructions: 'Synthetic test data only. Use args as CLI strings. Read-only.',
});
server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools: [{
  name: 'moff', description: 'Synthetic read-only Moff CLI fixture. Discover commands with schema and --help.',
  inputSchema: { type: 'object', properties: { args: { type: 'array', items: { type: 'string' } } },
    required: ['args'], additionalProperties: false },
}] }));
server.setRequestHandler(CallToolRequestSchema, async request => {
  const args = request.params.arguments?.args;
  let output;
  if (JSON.stringify(args) === '["schema"]') output = {
    schema: 'moff.schema/v1', command_prefixes: { default: ['outlook'] },
    commands: { mail: { list: true }, calendar: { folders: true, view: true } },
  };
  else if (Array.isArray(args) && args.includes('--help')) {
    return { content: [{ type: 'text', text:
      'outlook calendar view --folder-id <FOLDER_ID> --start <RFC3339> --end <RFC3339>' }] };
  } else if (JSON.stringify(args) === '["outlook","mail","list","--limit","2"]') output = {
    value: { Body: { ResponseMessages: { Items: [{ RootFolder: {
      TotalItemsInView: 3123, Items: [
        { ItemId: { Id: 'mail-alpha' }, Subject: 'Synthetic launch checklist', From: { Mailbox: { Name: 'Alice' } } },
        { ItemId: { Id: 'mail-beta' }, Subject: 'Synthetic budget review', From: { Mailbox: { Name: 'Bob' } } },
      ],
    } }] } } },
  };
  else if (JSON.stringify(args) === '["outlook","calendar","folders"]') output = {
    value: { Folders: [{ FolderId: { Id: 'calendar-real-fixture' }, DisplayName: 'Calendar', IsDefault: true }] },
  };
  else if (Array.isArray(args) && args.slice(0, 3).join(' ') === 'outlook calendar view' &&
    args[args.indexOf('--folder-id') + 1] === 'calendar-real-fixture' &&
    Number.isFinite(Date.parse(args[args.indexOf('--start') + 1])) &&
    Number.isFinite(Date.parse(args[args.indexOf('--end') + 1]))) output = {
    value: { Items: [{ ItemId: { Id: 'event-one' }, Subject: 'Synthetic Monday planning',
      Start: args[args.indexOf('--start') + 1], End: args[args.indexOf('--end') + 1] }] },
  };
  else return { isError: true, content: [{ type: 'text', text:
    'Invalid synthetic CLI call. Inspect --help; supported commands are outlook mail list --limit 2, outlook calendar folders, and outlook calendar view.' }] };
  return { content: [{ type: 'text', text: JSON.stringify(output) }] };
});
await server.connect(new StdioServerTransport());
