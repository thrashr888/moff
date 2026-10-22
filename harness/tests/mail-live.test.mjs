import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PersistentBotHarness } from '../dist/durable/harness.js';
import { DEFAULT_OLLAMA_MODEL } from '../dist/ollama/provider.js';

test('live mailbox request executes MCP, finishes, and supports a follow-up', { timeout: 480_000 }, async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-mail-live-'));
  const harness = new PersistentBotHarness({
    storagePath: join(directory, 'test.sqlite'),
    defaultModel: process.env.TEST_MODEL || DEFAULT_OLLAMA_MODEL,
  });
  try {
    await harness.init();
    const conversation = await harness.createConversation('Live verification');
    let characters = 0;
    await harness.sendMessage(conversation.id, 'Read my latest emails using moff.', {
      onDelta: delta => { characters += delta.length; },
      onTool: () => {},
    });
    const entries = await harness.getStorage().getEntries(conversation.id);
    const tools = entries.filter(entry => entry.role === 'tool');
    assert.ok(tools.length > 0, 'Model never called MCP');
    assert.ok(!tools.some(entry => entry.isError), 'MCP operation failed');
    console.log(`Mailbox turn finished: ${tools.length} tool results, ${characters} output characters`);
    // Verify numbered output cannot invent more records than the CLI returned.
    const answer = entries.filter(entry => entry.role === 'assistant').map(entry => entry.content || '').join('\n');
    const numbered = answer.match(/^\s*\d+[.)]\s/gm) || [];
    const records = [];
    function collect(value) {
      if (!value || typeof value !== 'object') return;
      if (typeof value.Subject === 'string' && value.ItemId) records.push(value);
      for (const child of Object.values(value)) {
        if (Array.isArray(child)) child.forEach(collect);
        else if (child && typeof child === 'object') collect(child);
      }
    }
    for (const tool of tools) collect(JSON.parse(tool.content));
    assert.ok(records.length > 0, 'No email records found in tool result');
    assert.ok(numbered.length <= records.length, 'Answer enumerated more emails than were returned');
    const subjects = records.map(record => record.Subject);
    assert.ok(subjects.some(subject => subject && answer.includes(subject)), 'Answer contains no source email subject');
    const normalize = text => text.toLowerCase().replace(/[\\*"'`]/g, '').replace(/\s+/g, ' ').trim();
    const listed = answer.split('\n').filter(line => /^\s*\d+[.)]\s/.test(line));
    assert.ok(listed.every(line => subjects.some(subject => normalize(line).includes(normalize(subject)))),
      'An enumerated email subject is not supported by the source');
    const source = tools.map(tool => tool.content).join('\n');
    const dates = answer.match(/\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}Z/g) || [];
    assert.ok(dates.every(date => source.includes(date)), 'Answer invented an ISO timestamp');
    const followup = await harness.sendMessage(conversation.id,
      'How many emails did the last tool result contain? Answer with only the number; do not read more emails.',
      { onDelta: () => {} });
    assert.equal(followup.trim(), String(records.length), 'Follow-up count differs from source record count');
    console.log(`Follow-up finished; ${records.length} source records checked without logging mailbox content`);
  } finally {
    await harness.close();
    await rm(directory, { recursive: true, force: true });
  }
});
