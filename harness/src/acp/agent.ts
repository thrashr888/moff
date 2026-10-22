import { agent, methods, PROTOCOL_VERSION, RequestError,
  type SessionConfigOption, type SessionNotification } from '@agentclientprotocol/sdk';
import type { PersistentBotHarness } from '../durable/harness.js';

export function createAcpAgent(harness: PersistentBotHarness, cwd: string) {
  const sessions = new Set<string>();
  const active = new Set<string>();
  const cancelled = new Set<string>();
  const config = async (id: string): Promise<SessionConfigOption[]> => [{
    id: 'model', name: 'Model', category: 'model', type: 'select',
    currentValue: (await harness.getConversation(id)).model,
    options: harness.availableModels().map(value => ({ value, name: value })),
  }];
  const attached = (id: string) => {
    if (!sessions.has(id)) throw new RequestError(-32602, `Load or create session ${id} first`);
  };
  return agent({ name: 'moff' })
    .onRequest('initialize', () => ({
      protocolVersion: PROTOCOL_VERSION,
      agentInfo: { name: 'moff', version: '0.1.0' },
      agentCapabilities: { loadSession: true, sessionCapabilities: { list: {} } },
    }))
    .onRequest('session/new', async ({ params }) => {
      if (params.mcpServers.length) throw new RequestError(-32602, 'Moff uses its own read-only MCP server');
      const conversation = await harness.createConversation();
      sessions.add(conversation.id);
      return { sessionId: conversation.id, configOptions: await config(conversation.id) };
    })
    .onRequest('session/list', async () => ({
      sessions: (await harness.listConversations()).map(session => ({
        sessionId: session.id, cwd, title: session.title,
      })),
    }))
    .onRequest('session/load', async ({ params, client }) => {
      if (params.mcpServers.length) throw new RequestError(-32602, 'External MCP servers are not supported');
      await harness.getConversation(params.sessionId);
      await harness.waitForIdle();
      for (const entry of await harness.getEntries(params.sessionId)) {
        if ((entry.role === 'user' || entry.role === 'assistant') && entry.content) {
          await client.notify(methods.client.session.update, {
            sessionId: params.sessionId, update: {
              sessionUpdate: entry.role === 'user' ? 'user_message_chunk' : 'agent_message_chunk',
              content: { type: 'text', text: entry.content },
              _meta: { 'moff/entryId': entry.id },
            },
          });
        }
      }
      sessions.add(params.sessionId);
      return { configOptions: await config(params.sessionId) };
    })
    .onRequest('session/set_config_option', async ({ params }) => {
      attached(params.sessionId);
      if (active.has(params.sessionId)) throw new RequestError(-32602, 'Wait for the current turn before changing models');
      if (params.configId !== 'model' || typeof params.value !== 'string') throw new RequestError(-32602, 'Unsupported configuration');
      if (!harness.availableModels().includes(params.value)) throw new RequestError(-32602, `Model not registered: ${params.value}`);
      await harness.setModel(params.sessionId, params.value);
      return { configOptions: await config(params.sessionId) };
    })
    .onRequest('session/prompt', async ({ params, client }) => {
      const id = params.sessionId;
      attached(id);
      if (active.has(id)) throw new RequestError(-32602, 'This session already has an active turn');
      if (params.prompt.some(block => block.type !== 'text')) throw new RequestError(-32602, 'Only text prompts are supported');
      const text = params.prompt.map(block => block.type === 'text' ? block.text : '').join('\n');
      if (!text.trim()) throw new RequestError(-32602, 'Prompt must not be empty');
      active.add(id);
      let notifications = Promise.resolve();
      const notify = (update: SessionNotification['update']) => {
        notifications = notifications.then(() => client.notify(methods.client.session.update, { sessionId: id, update }));
        // Keep the rejection observable at the end without an unhandled-rejection window.
        void notifications.catch(() => {});
      };
      try {
        await harness.sendMessage(id, text, {
          onDelta: delta => notify({ sessionUpdate: 'agent_message_chunk', content: { type: 'text', text: delta } }),
          onEvent: event => {
            if (event.type === 'tool_execution_start') notify({
              sessionUpdate: 'tool_call', toolCallId: event.toolCallId,
              title: `Reading with ${event.toolName}`, kind: 'read', status: 'in_progress',
            });
            if (event.type === 'tool_execution_end') notify({
              sessionUpdate: 'tool_call_update', toolCallId: event.toolCallId,
              status: event.entry?.model?.some(message => message.role === 'toolResult' && message.isError)
                ? 'failed' : event.entry ? 'completed' : 'failed',
            });
          },
        });
        await notifications;
        return { stopReason: cancelled.has(id) ? 'cancelled' : 'end_turn' };
      } catch (error) {
        await notifications;
        if (cancelled.has(id)) return { stopReason: 'cancelled' };
        throw new RequestError(-32603, error instanceof Error ? error.message : String(error));
      } finally {
        active.delete(id);
        cancelled.delete(id);
      }
    })
    .onNotification('session/cancel', async ({ params }) => {
      attached(params.sessionId);
      if (active.has(params.sessionId)) {
        cancelled.add(params.sessionId);
        await harness.abort(params.sessionId);
      }
    });
}
