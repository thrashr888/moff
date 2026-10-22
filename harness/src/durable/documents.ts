/**
 * Durable document definitions for persistent bot application state.
 * Documents are typed JSON state stored atomically alongside conversation transcripts.
 */

export interface DurableDocDefinition<T> {
  kind: string;
  version: number;
  scope: 'conversation' | 'global';
  initial: () => T;
}

export interface TodoItem {
  id: string;
  text: string;
  completed: boolean;
  createdAt: string;
}

export interface TodoDocState {
  items: TodoItem[];
}

export const TodoDoc: DurableDocDefinition<TodoDocState> = {
  kind: 'bot.todos',
  version: 1,
  scope: 'conversation',
  initial: () => ({ items: [] }),
};

export interface BotFact {
  key: string;
  value: string;
  updatedAt: string;
}

export interface BotMemoryDocState {
  facts: BotFact[];
  summary?: string;
}

export const BotMemoryDoc: DurableDocDefinition<BotMemoryDocState> = {
  kind: 'bot.memory',
  version: 1,
  scope: 'global',
  initial: () => ({ facts: [] }),
};

export interface BotSessionDocState {
  conversationId: string;
  title: string;
  model: string;
  createdAt: string;
  lastActiveAt: string;
  messageCount: number;
}

export const BotSessionDoc: DurableDocDefinition<BotSessionDocState> = {
  kind: 'bot.session_metadata',
  version: 1,
  scope: 'conversation',
  initial: () => ({
    conversationId: '',
    title: 'New Conversation',
    model: 'claude-3.5-sonnet',
    createdAt: new Date().toISOString(),
    lastActiveAt: new Date().toISOString(),
    messageCount: 0,
  }),
};
