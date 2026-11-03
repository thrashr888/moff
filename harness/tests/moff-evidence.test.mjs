import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { BACKGROUND_CONTEXT as context } from '@earendil-works/chord/context';
import { Type, createModels, fauxProvider, fauxAssistantMessage, fauxToolCall } from '@earendil-works/pi-ai';
import { Harness, createRegistry, defineExtension, defineTool } from '@earendil-works/pi-durable';
import { openNodeSqliteStorage } from '@earendil-works/pi-durable/storage/sqlite/node';
import { MoffEvidenceGuard, recordMoffEvidence } from '../dist/tools/moff-evidence.js';

test('Pi blocks invented IDs before execution and preserves discovered evidence across restart only within its conversation',
  { timeout: 20_000 }, async () => {
    const directory = await mkdtemp(join(tmpdir(), 'moff-evidence-'));
    const path = join(directory, 'pi.sqlite');
    const faux = fauxProvider();
    const models = createModels();
    models.setProvider(faux.provider);
    const model = faux.getModel();
    const agent = { model: { provider: model.provider, modelId: model.id } };
    const registry = createRegistry();
    let executions = 0;
    registry.install(MoffEvidenceGuard);
    registry.install(defineExtension({
      name: 'test-moff',
      tools: [defineTool({
        name: 'moff',
        parameters: Type.Object({ args: Type.Array(Type.String()) }),
        replay: 'safe',
        execute: async ({ args }, api, context) => {
          executions++;
          const text = args.includes('--help') ? 'Plain text help' : JSON.stringify({
            folders: [{ FolderId: { Id: 'real-folder' }, DisplayName: 'Calendar' }],
          });
          await recordMoffEvidence(text, api, context);
          return { output: [{ type: 'text', text }] };
        },
      })],
    }));
    let runtime;
    const open = async () => Harness.open(await openNodeSqliteStorage(path), { models, registry }, context);
    const send = async (conversation, args) => {
      faux.setResponses([
        fauxAssistantMessage(fauxToolCall('moff', { args }), { stopReason: 'toolUse' }),
        fauxAssistantMessage('DONE'),
      ]);
      const submission = await conversation.submit({ type: 'input', content: 'Test tool', whenBusy: 'reject' }, context);
      assert.equal((await submission.wait(context)).status, 'done');
      const entries = await conversation.entries({ order: 'descending' }, 10, undefined, context);
      return entries.items.flatMap(entry => entry.model || []).find(message => message.role === 'toolResult');
    };
    try {
      runtime = await open();
      let root = await runtime.root(context, { agent });
      let result = await send(root, ['outlook', 'calendar', 'view', '--folder-id', 'primary']);
      assert.equal(result.isError, true);
      assert.match(result.content[0].text, /not discovered/);
      assert.equal(executions, 0, 'Blocked ID must never reach Moff');
      result = await send(root, ['outlook', 'calendar', 'view', '--help']);
      assert.equal(result.isError, false, 'Help is exempt from ID checks and need not be JSON');
      await send(root, ['outlook', 'calendar', 'folders']);
      result = await send(root, ['outlook', 'calendar', 'view', '--folder-id=real-folder']);
      assert.equal(result.isError, false);
      assert.equal(executions, 3);
      await runtime.close(context);
      runtime = await open();
      root = await runtime.root(context, { agent });
      result = await send(root, ['outlook', 'calendar', 'view', '--folder-id', 'real-folder']);
      assert.equal(result.isError, false, 'Discovered IDs must survive reopen');
      const other = await runtime.createConversation({ ownership: { kind: 'ownerless' }, agent }, context);
      result = await send(other, ['outlook', 'calendar', 'view', '--folder-id', 'real-folder']);
      assert.equal(result.isError, true, 'Evidence must not leak across conversations');
      assert.equal(executions, 4);
    } finally {
      if (runtime) await runtime.close(context);
      await rm(directory, { recursive: true, force: true });
    }
  });
