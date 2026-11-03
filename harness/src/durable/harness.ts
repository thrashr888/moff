import { BACKGROUND_CONTEXT as context } from '@earendil-works/chord/context';
import { Type, type Models } from '@earendil-works/pi-ai';
import {
  Harness, createRegistry, defineExtension, defineTool, defineDoc, section, watchEvents,
  type Conversation, type ConversationId, type Cursor, type EntryId, type EntryRecord,
  type AgentEvent,
} from '@earendil-works/pi-durable';
import { openNodeSqliteStorage } from '@earendil-works/pi-durable/storage/sqlite/node';
import { DEFAULT_OLLAMA_MODEL, OllamaProvider } from '../ollama/provider.js';
import { MoffMcp } from '../tools/moff-mcp.js';
import { createBotModels, toChatMessage, type ModelProvider } from './models.js';
import type { StorageConversation, StorageEntry } from './storage.js';
import { LocalContext } from './context.js';
import { PartyIdentity } from './identity.js';
import { MoffEvidenceGuard, recordMoffEvidence } from '../tools/moff-evidence.js';

export type { ModelProvider } from './models.js';

export interface PersistentBotOptions {
  storagePath?: string;
  provider?: ModelProvider;
  models?: Models;
  providerId?: string;
  defaultModel?: string;
  enableMoff?: boolean;
  moffBinary?: string;
}

export interface ConversationTurnOptions {
  model?: string;
  requestId?: string;
  onDelta?: (delta: string) => void;
  onTool?: (message: string) => void;
  onEvent?: (event: AgentEvent) => void;
}

const Metadata = defineDoc({
  kind: 'moff.metadata', version: 1, scope: 'conversation', history: 'latest', fork: 'current',
  initial: () => ({ title: 'New Conversation', createdAt: new Date().toISOString() }),
});

export class PersistentBotHarness {
  private runtime?: Harness;
  private moff = new MoffMcp();
  private defaultModel: string;
  private providerId: string;
  private models: Models;
  private enableMoff: boolean;

  constructor(private options: PersistentBotOptions = {}) {
    this.defaultModel = options.defaultModel || DEFAULT_OLLAMA_MODEL;
    this.providerId = options.providerId || 'ollama';
    const provider = options.provider || new OllamaProvider();
    this.models = options.models || createBotModels(provider, this.providerId,
      [...new Set([this.defaultModel, 'qwen3:latest', 'llama3.2:latest'])]);
    this.enableMoff = options.enableMoff ?? (!options.models && provider instanceof OllamaProvider && provider.isLocal);
  }

  private get pi(): Harness {
    if (!this.runtime) throw new Error('Harness is not open');
    return this.runtime;
  }

  async init(): Promise<void> {
    if (this.runtime) throw new Error('Harness is already open');
    try {
      if (this.enableMoff) await this.moff.connect(this.options.moffBinary);
      const registry = createRegistry();
      registry.install(LocalContext);
      registry.install(PartyIdentity);
      registry.install(MoffEvidenceGuard);
      registry.install(defineExtension({
        name: 'moff',
        sections: [section('moff', () =>
          'You are Cipher, a local assistant running on Pi Durable with Moff tools. Match answer length to the request. ' +
          'Report only records present in tool results. Never invent entries or dates. Stop after the summary; do not repeat items. ' +
          (this.enableMoff
            ? 'You have read-only access through the Moff MCP server. You cannot send or modify mail. ' + this.moff.instructions
            : 'Mail tools are disabled for this provider. You cannot access emails.'), { tag: false })],
        tools: this.moff.tools.map(({ function: tool }) => defineTool({
          name: tool.name, description: tool.description || '',
          parameters: Type.Unsafe<Record<string, unknown>>(tool.parameters),
          replay: 'safe',
          execute: async (args, api, context) => {
            const text = await this.moff.call(tool.name, JSON.stringify(args));
            await recordMoffEvidence(text, api, context);
            return { output: [{ type: 'text', text }] };
          },
        })),
      }));
      this.runtime = await Harness.open(
        await openNodeSqliteStorage(this.options.storagePath || './pi-bot.sqlite'),
        { models: this.models, registry, settings: {
          toolExecution: 'sequential',
          retry: { enabled: false }, compaction: { reserveTokens: 4096, keepRecentTokens: 8192 },
        }, onReport: error => console.error('Pi Durable:', error) }, context);
      this.runtime.resume();
    } catch (error) {
      await this.moff.close();
      throw error;
    }
  }

  private agent(model = this.defaultModel) {
    if (!this.models.getModel(this.providerId, model)) throw new Error(`Model not registered: ${this.providerId}/${model}`);
    return { model: { provider: this.providerId, modelId: model } };
  }

  async resume(): Promise<number> {
    const pending = (await this.pi.inspect(context)).tasks.length;
    this.pi.resume();
    return pending;
  }

  async rootConversation(): Promise<StorageConversation> {
    const root = await this.pi.root(context, { agent: this.agent() });
    return this.describe(root);
  }

  async createConversation(title?: string, model?: string): Promise<StorageConversation> {
    const conversation = await this.pi.createConversation({
      ownership: { kind: 'ownerless' }, agent: this.agent(model),
      init: async (tx, id) => { (await tx.doc(Metadata, id)).title = title || 'New Conversation'; },
    }, context);
    return this.describe(conversation);
  }

  async forkConversation(parentId: string, forkEntryId?: string): Promise<StorageConversation> {
    const parent = await this.conversation(parentId);
    let at: EntryId;
    if (forkEntryId !== undefined) {
      if (!/^[1-9]\d*$/.test(forkEntryId) || !Number.isSafeInteger(Number(forkEntryId))) {
        throw new Error(`Invalid fork entry ID: ${forkEntryId}`);
      }
      at = Number(forkEntryId) as EntryId;
    } else {
      const latest = (await parent.entries({}, 1, undefined, context)).items[0];
      if (!latest) {
        const info = await this.describe(parent);
        return this.createConversation(`Fork of ${info.title}`, info.model);
      }
      at = latest.id;
    }
    const info = await this.describe(parent);
    const fork = await parent.fork(at, {
      ownership: { kind: 'ownerless' },
      init: async (tx, id) => {
        const metadata = await tx.doc(Metadata, id);
        metadata.title = `Fork of ${info.title}`;
        metadata.createdAt = new Date().toISOString();
      },
    }, context);
    return this.describe(fork);
  }

  private async conversation(id: string): Promise<Conversation> {
    if (!/^[1-9]\d*$/.test(id) || !Number.isSafeInteger(Number(id))) throw new Error(`Invalid conversation ID: ${id}`);
    // Decimal IDs are decoded only at this external boundary; Pi owns their allocation.
    const conversation = await this.pi.conversation(Number(id) as ConversationId, context);
    if (!conversation) throw new Error(`Conversation ${id} not found`);
    return conversation;
  }

  private async describe(conversation: Conversation): Promise<StorageConversation> {
    const record = await this.pi.commit(tx => tx.conversation(conversation.id), context);
    const meta = await conversation.commit(async tx => ({ ...await tx.doc(Metadata, conversation.id) }), context);
    const agent = await conversation.agent(context);
    return { id: String(conversation.id), title: meta.title, model: agent.model?.modelId || this.defaultModel,
      createdAt: meta.createdAt, updatedAt: meta.createdAt,
      parentId: record?.parent ? String(record.parent.conversationId) : undefined,
      forkEntryId: record?.parent ? String(record.parent.at) : undefined };
  }

  async getConversation(id: string): Promise<StorageConversation> {
    return this.describe(await this.conversation(id));
  }

  async setModel(id: string, model: string): Promise<void> {
    await (await this.conversation(id)).configure(this.agent(model), context);
  }

  availableModels(): string[] { return this.models.getModels(this.providerId).map(model => model.id); }

  async listConversations(): Promise<StorageConversation[]> {
    const conversations: StorageConversation[] = [];
    let cursor: Cursor | undefined;
    do {
      const page = await this.pi.commit(tx => tx.scanConversations({}, 100, cursor), context);
      for (const record of page.items) {
        const conversation = await this.pi.conversation(record.id, context);
        if (conversation) conversations.push(await this.describe(conversation));
      }
      cursor = page.next;
    } while (cursor);
    return conversations;
  }

  async getEntries(id: string): Promise<StorageEntry[]> {
    const conversation = await this.conversation(id);
    const entries: StorageEntry[] = [];
    let cursor: Cursor | undefined;
    do {
      const page = await conversation.entries({ order: 'ascending' }, 100, cursor, context);
      for (const entry of page.items) {
        for (const message of entry.model || []) {
          const chat = toChatMessage(message);
          entries.push({ id: String(entry.id), conversationId: id, role: chat.role, content: chat.content,
            toolCalls: chat.tool_calls, toolCallId: chat.tool_call_id,
            isError: message.role === 'toolResult' ? message.isError : undefined,
            createdAt: new Date(message.timestamp).toISOString() });
        }
      }
      cursor = page.next;
    } while (cursor);
    return entries;
  }

  getStorage() { return { getEntries: (id: string) => this.getEntries(id) }; }

  async sendMessage(id: string, userText: string, options: ConversationTurnOptions = {}): Promise<string> {
    const conversation = await this.conversation(id);
    if (options.model) await conversation.configure(this.agent(options.model), context);
    const events = await watchEvents(this.pi, conversation.id, context);
    let streamed = '';
    let currentText = '';
    const blocks = new Map<number, string>();
    const emit = (value: string) => {
      const delta = value.startsWith(currentText) ? value.slice(currentText.length) : '';
      if (delta) { streamed += delta; options.onDelta?.(delta); }
      currentText = value;
    };
    const entryText = (entry: EntryRecord) => (entry.model || []).flatMap(message =>
      message.role === 'assistant' ? message.content.filter(block => block.type === 'text').map(block => block.text) : []).join('');
    const setBlocks = (content: readonly { type: string; text?: string }[]) => {
      blocks.clear();
      content.forEach((block, index) => {
        if (block.type === 'text') blocks.set(index, block.text || '');
      });
    };
    const emitBlocks = () => emit([...blocks].sort(([a], [b]) => a - b).map(([, text]) => text).join(''));
    events.start(async batch => {
      for (const event of batch) {
        options.onEvent?.(event);
        if (event.type === 'message_start' && event.message.role === 'assistant') {
          currentText = '';
          setBlocks(event.message.content);
          emitBlocks();
        }
        if (event.type === 'message_update') {
          for (const change of event.changes) {
            if (change.type === 'text_delta') blocks.set(change.contentIndex, (blocks.get(change.contentIndex) || '') + change.delta);
            if (change.type === 'text_start' || change.type === 'block') {
              if (change.block.type === 'text') blocks.set(change.contentIndex, change.block.text);
              else blocks.delete(change.contentIndex);
            }
            if (change.type === 'message') setBlocks(change.message.content);
          }
          emitBlocks();
        }
        if (event.type === 'message_end' && event.entry.kind === 'pi.assistant') emit(entryText(event.entry));
        if (event.type === 'tool_execution_start') options.onTool?.(`Reading with ${event.toolName}...`);
      }
    });
    try {
      const submission = await conversation.submit({
        type: 'input', content: userText, requestId: options.requestId, whenBusy: 'reject',
      }, context);
      const settled = await submission.wait(context);
      if (settled.status !== 'done' || settled.type !== 'input') {
        const entries = await this.getEntries(id);
        const detail = (await conversation.context(context)).entries.flatMap(entry => entry.model || [])
          .filter(message => message.role === 'assistant' && message.errorMessage)
          .at(-1);
        throw new Error(detail?.role === 'assistant' ? detail.errorMessage
          : `Input was not answered: ${settled.reason}. Transcript entries: ${entries.length}`);
      }
      const answer = await conversation.commit(tx => tx.entry(settled.answer), context);
      const final = answer?.model?.find(message => message.role === 'assistant');
      if (final?.stopReason === 'length') throw new Error('The model stopped at its output limit. The answer is incomplete.');
      return answer ? entryText(answer) : streamed;
    } finally {
      await events.stop();
    }
  }

  async abort(id: string): Promise<void> { await (await this.conversation(id)).abort(context); }
  async compact(id: string): Promise<void> { await (await this.conversation(id)).compact(undefined, context); }
  async waitForIdle(): Promise<void> { await this.pi.waitForIdle(context); }

  async close(): Promise<void> {
    const runtime = this.runtime;
    this.runtime = undefined;
    try { if (runtime) await runtime.close(context); }
    finally { await this.moff.close(); }
  }
}
