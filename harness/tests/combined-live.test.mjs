import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { PersistentBotHarness } from '../dist/durable/harness.js';
import { DEFAULT_OLLAMA_MODEL } from '../dist/ollama/provider.js';

test('local model handles the reported combined two-email and Monday-calendar request', { timeout: 480_000 }, async t => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-combined-live-'));
  const harness = new PersistentBotHarness({ storagePath: join(directory, 'test.sqlite'),
    defaultModel: process.env.TEST_MODEL || DEFAULT_OLLAMA_MODEL });
  let conversation;
  const cancel = () => {
    if (conversation) {
      void harness.abort(conversation.id).catch(error => console.error('Live-test cancellation failed:', error));
    }
  };
  try {
    await harness.init();
    conversation = await harness.createConversation();
    const requestedAt = new Date();
    t.signal.addEventListener('abort', cancel, { once: true });
    const answer = await harness.sendMessage(conversation.id,
      "what's my latest 2 emails, and my calendar for monday?", {
        onDelta: () => {},
        onTool: message => console.log(message),
      });
    const entries = await harness.getEntries(conversation.id);
    const calls = entries.flatMap(entry => entry.toolCalls || []);
    const moff = calls.filter(call => call.function.name === 'moff').map(call => {
      const args = JSON.parse(call.function.arguments);
      assert.ok(args && !Array.isArray(args) && Array.isArray(args.args), 'Moff arguments must be an object with args');
      return args.args;
    });
    assert.ok(moff.some(args => args.slice(0, 3).join(' ') === 'outlook mail list' &&
      args[args.indexOf('--limit') + 1] === '2'), 'Did not request exactly two emails');
    const results = entries.filter(entry => entry.role === 'tool');
    const successfulCalls = new Set();
    const emails = new Map();
    const calendar = [];
    function collect(value, records) {
      if (!value || typeof value !== 'object') return;
      if (typeof value.Subject === 'string' && value.ItemId) records.push(value);
      for (const child of Object.values(value)) {
        collect(child, records);
      }
    }
    for (const result of results) {
      const call = calls.find(call => call.id === result.toolCallId);
      if (result.isError) continue;
      if (call?.function.name === 'moff' && JSON.parse(call.function.arguments).args.includes('--help')) continue;
      const output = JSON.parse(result.content);
      successfulCalls.add(result.toolCallId);
      if (call?.function.name !== 'moff') continue;
      const args = JSON.parse(call.function.arguments).args;
      if (args.slice(0, 3).join(' ') === 'outlook mail list') {
        assert.equal(args[args.indexOf('--limit') + 1], '2', 'Mail list must request only two emails');
        const page = [];
        collect(output, page);
        assert.equal(page.length, 2, 'Mail tool must return two email records');
        for (const email of page) emails.set(JSON.stringify(email.ItemId), email);
      }
      if (args.slice(0, 3).join(' ') === 'outlook calendar view') {
        collect(output, calendar);
      }
    }
    const records = [...emails.values()];
    const normalize = text => text.toLowerCase().replace(/[\\*"'`]/g, '').replace(/\s+/g, ' ').trim();
    console.log(JSON.stringify({ model: process.env.TEST_MODEL || DEFAULT_OLLAMA_MODEL, calls: calls.map(call => ({
      name: call.function.name,
      command: call.function.name === 'moff' ? JSON.parse(call.function.arguments).args.slice(0, 3).join(' ') : undefined,
      succeeded: successfulCalls.has(call.id),
    })), answerCharacters: answer.length,
      supportedSubjects: records.filter(record => normalize(answer).includes(normalize(record.Subject))).length }));
    assert.equal(records.length, 2, 'Must read exactly two distinct source emails');
    assert.ok(records.every(record => normalize(answer).includes(normalize(record.Subject))), 'Answer must include both source email subjects');
    assert.ok(calls.some(call => call.function.name === 'current_time'), 'Relative dates require current clock context');
    assert.ok(moff.some(args => args.join(' ') === 'outlook calendar folders'), 'Calendar must discover real folders');
    const views = calls.filter(call => call.function.name === 'moff' && successfulCalls.has(call.id))
      .map(call => JSON.parse(call.function.arguments).args)
      .filter(args => args.slice(0, 3).join(' ') === 'outlook calendar view' && !args.includes('--help'));
    assert.ok(views.length > 0, 'Must successfully query the calendar, not merely report a model/tool failure');
    const monday = new Date(requestedAt);
    const days = (8 - monday.getDay()) % 7 || 7;
    monday.setDate(monday.getDate() + days);
    monday.setHours(0, 0, 0, 0);
    for (const args of views) {
      const start = new Date(args[args.indexOf('--start') + 1]);
      assert.equal(start.getTime(), monday.getTime(), 'Calendar must use the upcoming Monday in the local timezone');
      const tuesday = new Date(monday);
      tuesday.setDate(tuesday.getDate() + 1);
      assert.equal(new Date(args[args.indexOf('--end') + 1]).getTime(), tuesday.getTime(),
        'Calendar range must end at the next local midnight');
      assert.notEqual(args[args.indexOf('--folder-id') + 1], 'primary', 'Folder ID must not be invented');
    }
    if (calendar.length) {
      assert.ok(calendar.some(record => normalize(answer).includes(normalize(record.Subject))),
        'Calendar answer must include a source event subject');
    }
    console.log(`Combined request completed: ${calls.length} calls, ${answer.length} answer characters; source content not logged`);
  } finally {
    t.signal.removeEventListener('abort', cancel);
    try { await harness.close(); }
    finally { await rm(directory, { recursive: true, force: true }); }
  }
});
