import { test } from 'node:test';
import assert from 'node:assert/strict';
import { OllamaProvider } from '../dist/ollama/provider.js';

for (const streaming of [false, true]) {
  test(`Ollama normalizes null tool-call content (stream=${streaming})`, async t => {
    const call = { id: 'call-1', type: 'function', function: { name: 'moff', arguments: '{"args":["schema"]}' } };
    const messages = [
      { role: 'user', content: 'Read my latest emails' },
      { role: 'assistant', content: null, tool_calls: [call] },
      { role: 'tool', content: '{"schema":"moff.schema/v1"}', tool_call_id: 'call-1' },
    ];
    t.mock.method(globalThis, 'fetch', async (_url, options) => {
      const body = JSON.parse(options.body);
      assert.equal(body.model, 'qwen3.5:9b');
      assert.equal(body.reasoning_effort, 'none');
      assert.equal(body.messages[1].content, '');
      assert.deepEqual(body.messages[1].tool_calls, [call]);
      assert.deepEqual(body.messages[2], messages[2]);
      assert.ok(body.messages.every(message => typeof message.content === 'string'));
      if (streaming) {
        return new Response('data: {"id":"x","choices":[{"delta":{"content":"Done"},"finish_reason":"stop"}]}\n\ndata: [DONE]\n\n');
      }

      return Response.json({
        id: 'x', model: 'qwen3:latest',
        choices: [{ message: { role: 'assistant', content: 'Done' }, finish_reason: 'stop' }],
      });
    });
    const provider = new OllamaProvider();
    if (streaming) {
      let text = '';
      for await (const delta of provider.stream({ messages })) text += delta.content || '';
      assert.equal(text, 'Done');
    } else {
      assert.equal((await provider.complete({ messages })).message.content, 'Done');
    }
    assert.equal(messages[1].content, null, 'Stored messages must not be mutated');
  });
}

test('reasoning-only/empty completions fail explicitly in both modes', async t => {
  const provider = new OllamaProvider();
  t.mock.method(globalThis, 'fetch', async () => Response.json({
    id: 'x', model: 'qwen3:latest', choices: [{
      message: { role: 'assistant', content: '', reasoning: 'I should call a tool' }, finish_reason: 'stop',
    }],
  }));
  await assert.rejects(provider.complete({ messages: [] }), /neither answer text nor tool calls/);
  t.mock.restoreAll();
  t.mock.method(globalThis, 'fetch', async () => new Response(
    'data: {"id":"x","choices":[{"delta":{},"finish_reason":"stop"}]}\n\ndata: [DONE]\n\n',
  ));
  await assert.rejects(async () => {
    for await (const delta of provider.stream({ messages: [] })) void delta;
  }, /neither answer text nor tool calls/);
});

test('Qwen3.5 uses thinking-off configuration without applying it to unrelated model families', async t => {
  t.mock.method(globalThis, 'fetch', async (_url, options) => {
    const body = JSON.parse(options.body);
    assert.equal(body.reasoning_effort, body.model === 'qwen3.5:9b' ? 'none' : undefined);
    return Response.json({
      id: 'x', model: body.model,
      choices: [{ message: { role: 'assistant', content: 'Done' }, finish_reason: 'stop' }],
    });
  });
  for (const model of ['qwen3.5:9b', 'hf.co/unsloth/granite-4.1-8b-GGUF:Q4_K_M']) {
    assert.equal((await new OllamaProvider().complete({ model, messages: [] })).message.content, 'Done');
  }
});
