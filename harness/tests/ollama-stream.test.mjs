import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OllamaProvider } from '../dist/ollama/provider.js';

test('Ollama stream preserves fragmented tool calls and text', async t => {
  const frames = [
    { id: 'chat-1', choices: [{ delta: { tool_calls: [{ index: 0, id: 'call-1', function: { name: 'moff_mail_list', arguments: '{"lim' } }] } }] },
    { id: 'chat-1', choices: [{ delta: { tool_calls: [{ index: 0, function: { arguments: 'it":5}' } }] } }] },
    { id: 'chat-1', choices: [{ delta: { content: 'Ready' }, finish_reason: 'tool_calls' }] },
  ];
  const data = frames.map(frame => `data: ${JSON.stringify(frame)}\r\n\r\n`).join('') + 'data: [DONE]\n\n';
  t.mock.method(globalThis, 'fetch', async () => new Response(new ReadableStream({
    start(controller) {
      const bytes = new TextEncoder().encode(data);
      for (let i = 0; i < bytes.length; i += 7) controller.enqueue(bytes.slice(i, i + 7));
      controller.close();
    },
  })));
  const stream = new OllamaProvider().stream({ messages: [] });
  let text = '';
  let response;
  while (true) {
    const next = await stream.next();
    if (next.done) { response = next.value; break; }
    text += next.value.content || '';
  }
  assert.equal(text, 'Ready');
  assert.equal(response.finish_reason, 'tool_calls');
  assert.deepEqual(response.message.tool_calls, [
    { id: 'call-1', type: 'function', function: { name: 'moff_mail_list', arguments: '{"limit":5}' } },
  ]);
});

test('Ollama rejects malformed and truncated streams instead of silently succeeding', async t => {
  for (const data of ['data: not-json\n', 'data: {"id":"x","choices":[{"delta":{"content":"partial"}}]}\n']) {
    t.mock.method(globalThis, 'fetch', async () => new Response(data));
    const stream = new OllamaProvider().stream({ messages: [] });
    await assert.rejects(async () => { for await (const delta of stream) { void delta; } });
    t.mock.restoreAll();
  }
});

test('Ollama leaves answer length to the model and preserves the length finish reason', async t => {
  let body;
  t.mock.method(globalThis, 'fetch', async (_url, options) => {
    body = JSON.parse(options.body);
    return new Response('data: {"id":"bounded","choices":[{"delta":{"content":"short"},"finish_reason":"length"}]}\n\ndata: [DONE]\n\n');
  });

  const stream = new OllamaProvider().stream({ messages: [] });
  await stream.next();
  const final = await stream.next();
  assert.equal(body.max_tokens, undefined);
  assert.equal(final.value.finish_reason, 'length');
});

test('Ollama cancels a repeating numbered entry stream', async t => {
  let canceled = false;
  t.mock.method(globalThis, 'fetch', async () => new Response(new ReadableStream({
    start(controller) {
      for (let i = 1; i <= 7; i++) {
        const frame = { id: 'loop', choices: [{ delta: { content: `${i}. An identical fabricated email entry from the same sender and timestamp\n` } }] };
        controller.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(frame)}\n\n`));
      }
    },
    cancel() { canceled = true; },
  })));
  const stream = new OllamaProvider().stream({ messages: [] });
  await assert.rejects(async () => { for await (const delta of stream) { void delta; } }, /repetitive model output/);
  assert.equal(canceled, true);
});

test('stream is canceled before hundreds of timestamp-variant entries can print', async t => {
  let canceled = false;
  let entry = 0;
  t.mock.method(globalThis, 'fetch', async () => new Response(new ReadableStream({
    pull(controller) {
      entry++;
      const time = `2026-10-12T22:${String(entry % 60).padStart(2, '0')}:00Z`;
      const frame = { id: 'varying-loop', choices: [{
        delta: { content: `${entry}. "George Simmons" from George Simmons, sent on ${time}\n` },
      }] };
      controller.enqueue(new TextEncoder().encode(`data: ${JSON.stringify(frame)}\n\n`));
    },
    cancel() { canceled = true; },
  })));
  const stream = new OllamaProvider().stream({
    messages: [{ role: 'tool', content: '{"messages":[{"subject":"George Simmons","date":"2026-10-12T22:01:00Z"}]}' }],
  });
  let emitted = 0;
  await assert.rejects(async () => {
    for await (const delta of stream) {
      if (delta.content) emitted++;
    }
  }, /unsupported timestamps/);
  assert.ok(emitted < 10, `Emitted ${emitted} entries`);
  assert.equal(canceled, true);
});
