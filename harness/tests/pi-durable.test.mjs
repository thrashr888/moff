import { test } from 'node:test';
import assert from 'node:assert/strict';
import { fork } from 'node:child_process';
import { once } from 'node:events';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { createModels, fauxProvider, fauxAssistantMessage } from '@earendil-works/pi-ai';
import { PersistentBotHarness } from '../dist/durable/harness.js';
import { setTimeout as delay } from 'node:timers/promises';
import { toChatMessage } from '../dist/durable/models.js';

test('model adapter preserves explicit tool failure semantics', () => {
  const base = { role: 'toolResult', toolCallId: 'failed', toolName: 'moff',
    content: [{ type: 'text', text: 'Unknown folder' }], timestamp: Date.now() };
  assert.match(toChatMessage({ ...base, isError: true }).content, /^Tool failed;.*\nUnknown folder$/);
  assert.equal(toChatMessage({ ...base, isError: false }).content, 'Unknown folder');
});

test('Pi serializes a batch of native MCP calls instead of queueing overlapping requests', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-pi-sequential-'));
  const events = [];
  const provider = {
    async complete() { throw new Error('Unexpected nonstreaming call'); },
    async *stream(request) {
      if (!request.messages.some(message => message.role === 'tool')) return {
        id: 'batch', model: 'test', finish_reason: 'tool_calls',
        message: { role: 'assistant', content: '', tool_calls: [
          { id: 'first', type: 'function', function: { name: 'moff', arguments: '{"args":["schema"]}' } },
          { id: 'second', type: 'function', function: { name: 'moff', arguments: '{"args":["outlook","calendar","view","--help"]}' } },
        ] },
      };
      yield { content: 'DONE' };
      return { id: 'done', model: 'test', finish_reason: 'stop',
        message: { role: 'assistant', content: 'DONE' } };
    },
  };
  const harness = new PersistentBotHarness({ provider, defaultModel: 'test',
    storagePath: join(directory, 'pi.sqlite'), enableMoff: true });
  try {
    await harness.init();
    const root = await harness.rootConversation();
    await harness.sendMessage(root.id, 'Read schema and calendar help', {
      onEvent: event => {
        if (event.type === 'tool_execution_start' || event.type === 'tool_execution_end') {
          events.push({ type: event.type, id: event.toolCallId });
        }
      },
    });
    assert.deepEqual(events, [
      { type: 'tool_execution_start', id: 'first' },
      { type: 'tool_execution_end', id: 'first' },
      { type: 'tool_execution_start', id: 'second' },
      { type: 'tool_execution_end', id: 'second' },
    ]);
    const results = (await harness.getEntries(root.id)).filter(entry => entry.role === 'tool');
    assert.equal(results.length, 2);
    assert.ok(results.every(entry => !entry.isError));
  } finally {
    await harness.close();
    await rm(directory, { recursive: true, force: true });
  }
});

test('committed streaming snapshots preserve the first chunk and all later deltas', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-pi-stream-'));
  const provider = {
    async complete() { throw new Error('Unexpected nonstreaming call'); },
    async *stream(request) {
      assert.ok(request.messages.some(message => message.role === 'system' &&
        message.content.includes('Your chosen name is Cipher') &&
        message.content.includes('Ranger (Gloom Stalker)')), 'Bot identity must be available in real Pi context');
      for (const content of ['NATIVE', '_ACP', '_OK']) {
        yield { content };
        await delay(160);
      }
      return { id: 'stream', model: 'test', finish_reason: 'stop',
        message: { role: 'assistant', content: 'NATIVE_ACP_OK' } };
    },
  };
  const harness = new PersistentBotHarness({ provider, defaultModel: 'test',
    storagePath: join(directory, 'pi.sqlite'), enableMoff: false });
  try {
    await harness.init();
    const root = await harness.rootConversation();
    let streamed = '';
    assert.equal(await harness.sendMessage(root.id, 'Hello', { onDelta: delta => { streamed += delta; } }), 'NATIVE_ACP_OK');
    assert.equal(streamed, 'NATIVE_ACP_OK');
  } finally {
    await harness.close();
    await rm(directory, { recursive: true, force: true });
  }
});

test('Pi Durable reopens the same root and preserves multi-turn context', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-pi-'));
  const faux = fauxProvider();
  const models = createModels();
  models.setProvider(faux.provider);
  const model = faux.getModel();
  const options = { models, providerId: model.provider, defaultModel: model.id,
    enableMoff: false, storagePath: join(directory, 'pi.sqlite') };
  let harness = new PersistentBotHarness(options);
  try {
    await harness.init();
    const root = await harness.rootConversation();
    faux.setResponses([fauxAssistantMessage('FIRST')]);
    let first = '';
    assert.equal(await harness.sendMessage(root.id, 'Remember the word penguin.', {
      requestId: 'first', onDelta: delta => { first += delta; },
    }), 'FIRST');
    assert.equal(first, 'FIRST', 'Committed text must be streamed exactly once');
    await harness.close();
    harness = new PersistentBotHarness(options);
    await harness.init();
    assert.equal((await harness.rootConversation()).id, root.id);
    faux.setResponses([(context) => {
      assert.ok(context.messages.some(message => message.role === 'user' && message.content === 'Remember the word penguin.'));
      return fauxAssistantMessage('penguin');
    }]);
    assert.equal(await harness.sendMessage(root.id, 'What was the word?'), 'penguin');
    const before = faux.state.callCount;
    assert.equal(await harness.sendMessage(root.id, 'Remember the word penguin.', { requestId: 'first' }), 'FIRST');
    assert.equal(faux.state.callCount, before, 'A retried request must not generate again');
    assert.equal((await harness.getEntries(root.id)).filter(entry => entry.role === 'user').length, 2);
    const forkAt = (await harness.getEntries(root.id)).filter(entry => entry.role === 'assistant')[0].id;
    const fork = await harness.forkConversation(root.id, forkAt);
    assert.equal(fork.parentId, root.id);
    assert.equal(fork.forkEntryId, forkAt);
    const inherited = await harness.getEntries(fork.id);
    assert.equal(inherited.filter(entry => entry.role === 'user').length, 1);
    assert.equal(inherited.filter(entry => entry.role === 'assistant').at(-1).content, 'FIRST');
  } finally {
    await harness.close();
    await rm(directory, { recursive: true, force: true });
  }
});

test('a killed process resumes its pending Pi generation without repeating a committed MCP call',
  { timeout: 20_000 }, async () => {
    const directory = await mkdtemp(join(tmpdir(), 'moff-pi-crash-'));
    const workerURL = new URL('./fixtures/pi-worker.mjs', import.meta.url);
    const start = mode => fork(workerURL, [], {
      env: { ...process.env, PI_TEST_PATH: join(directory, 'pi.sqlite'), PI_TEST_MODE: mode },
      stdio: ['ignore', 'ignore', 'pipe', 'ipc'],
    });
    let child;
    try {
      child = start('interrupt');
      let errors = '';
      child.stderr.on('data', data => { errors += data; });
      const [ready] = await Promise.race([
        once(child, 'message'),
        once(child, 'exit').then(([code]) => { throw new Error(`Worker exited ${code}: ${errors}`); }),
      ]);
      assert.equal(ready.type, 'pending');
      const exit = once(child, 'exit');
      child.kill('SIGKILL');
      await exit;
      child = start('resume');
      child.stderr.on('data', data => { errors += data; });
      const resultPromise = once(child, 'message');
      const exitPromise = once(child, 'exit');
      const [result] = await Promise.race([
        resultPromise,
        exitPromise.then(([code]) => { throw new Error(`Recovery worker exited ${code}: ${errors}`); }),
      ]);
      assert.deepEqual(result, { type: 'recovered', sameRoot: true, users: 1, tools: 1, answer: 'RECOVERED' });
      assert.equal((await exitPromise)[0], 0, errors);
    } finally {
      if (child && child.exitCode === null && child.signalCode === null) {
        const exit = once(child, 'exit');
        child.kill('SIGKILL');
        await exit;
      }
      await rm(directory, { recursive: true, force: true });
    }
  });
