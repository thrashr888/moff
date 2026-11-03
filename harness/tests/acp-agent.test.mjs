import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { client } from '@agentclientprotocol/sdk';
import { createModels, fauxProvider, fauxAssistantMessage } from '@earendil-works/pi-ai';
import { PersistentBotHarness } from '../dist/durable/harness.js';
import { createAcpAgent } from '../dist/acp/agent.js';

test('ACP SDK negotiates, creates, streams, lists, reloads and rejects unsupported input', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-acp-'));
  const faux = fauxProvider();
  const model = faux.getModel();
  const models = createModels();
  models.setProvider(faux.provider);
  const harness = new PersistentBotHarness({ models, providerId: model.provider,
    defaultModel: model.id, storagePath: join(directory, 'pi.sqlite'), enableMoff: false });
  const updates = [];
  let connection;
  try {
    await harness.init();
    connection = client().onNotification('session/update', ({ params }) => { updates.push(params); })
      .connect(createAcpAgent(harness, process.cwd()));
    const request = (method, params) => connection.agent.request(method, params);
    const initialized = await request('initialize', { protocolVersion: 1, clientCapabilities: {} });
    assert.equal(initialized.protocolVersion, 1);
    assert.equal(initialized.agentCapabilities.loadSession, true);
    const params = { cwd: process.cwd(), mcpServers: [] };
    const session = await request('session/new', params);
    assert.equal(session.configOptions[0].currentValue, model.id);
    faux.setResponses([fauxAssistantMessage('ACP_OK')]);
    assert.deepEqual(await request('session/prompt', { sessionId: session.sessionId,
      prompt: [{ type: 'text', text: 'Hello' }] }), { stopReason: 'end_turn' });
    assert.equal(updates.filter(item => item.update.sessionUpdate === 'agent_message_chunk')
      .map(item => item.update.content.text).join(''), 'ACP_OK');
    assert.ok(updates.every(item => item.sessionId === session.sessionId));
    assert.ok((await request('session/list', {})).sessions.some(item => item.sessionId === session.sessionId));
    updates.length = 0;
    await request('session/load', { ...params, sessionId: session.sessionId });
    assert.deepEqual(updates.map(item => item.update.content.text), ['Hello', 'ACP_OK']);
    await assert.rejects(request('session/prompt', { sessionId: session.sessionId,
      prompt: [{ type: 'image', data: 'AA==', mimeType: 'image/png' }] }), /Only text prompts/);
    await assert.rejects(request('session/set_config_option', {
      sessionId: session.sessionId, configId: 'model', value: 'unavailable',
    }), /Model not registered/);
  } finally {
    await connection?.close();
    await harness.close();
    await rm(directory, { recursive: true, force: true });
  }
});

test('ACP cancellation aborts the underlying durable generation', { timeout: 10_000 }, async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-acp-cancel-'));
  let ready;
  const started = new Promise(resolve => { ready = resolve; });
  const provider = {
    async complete() { throw new Error('Unexpected nonstreaming call'); },
    async *stream(request) {
      ready();
      await new Promise((_, reject) => {
        request.signal.addEventListener('abort', () => reject(new Error('Cancelled')), { once: true });
      });
      throw new Error('Should not finish');
    },
  };
  const harness = new PersistentBotHarness({ provider, defaultModel: 'test',
    storagePath: join(directory, 'pi.sqlite'), enableMoff: false });
  let connection;
  try {
    await harness.init();
    connection = client().onNotification('session/update', () => {})
      .connect(createAcpAgent(harness, process.cwd()));
    const session = await connection.agent.request('session/new', { cwd: process.cwd(), mcpServers: [] });
    const result = connection.agent.request('session/prompt', {
      sessionId: session.sessionId, prompt: [{ type: 'text', text: 'Wait' }],
    });
    await started;
    await connection.agent.notify('session/cancel', { sessionId: session.sessionId });
    assert.equal((await result).stopReason, 'cancelled');
    await harness.waitForIdle();
  } finally {
    await connection?.close();
    await harness.close();
    await rm(directory, { recursive: true, force: true });
  }
});
