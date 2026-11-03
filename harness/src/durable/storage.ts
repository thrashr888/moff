import * as path from 'node:path';
import * as fs from 'node:fs';

export interface StorageEntry {
  id: string;
  conversationId: string;
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string | null;
  toolCalls?: Array<{
    id: string;
    type: 'function';
    function: { name: string; arguments: string };
  }>;
  toolCallId?: string;
  isError?: boolean;
  createdAt: string;
}

export interface StorageConversation {
  id: string;
  parentId?: string;
  forkEntryId?: string;
  title: string;
  model: string;
  createdAt: string;
  updatedAt: string;
}

export interface StorageTask {
  id: string;
  conversationId: string;
  name: string;
  status: 'pending' | 'running' | 'completed' | 'failed' | 'aborted';
  checkpoint?: string;
  createdAt: string;
  updatedAt: string;
}

export interface StorageDocument<T = unknown> {
  kind: string;
  scopeId: string;
  version: number;
  data: T;
  updatedAt: string;
}

/**
 * Durable storage backend for conversations, transcripts, tasks, and documents.
 * Compatible with Node.js built-in sqlite (node:sqlite) or file-based persistent fallback.
 */
export class DurableStorage {
  private dbPath: string;
  private db: any = null;
  private memoryStore: {
    conversations: Map<string, StorageConversation>;
    entries: Map<string, StorageEntry[]>;
    tasks: Map<string, StorageTask>;
    documents: Map<string, StorageDocument>;
  };

  constructor(dbPath: string = './bot-storage.sqlite') {
    this.dbPath = path.resolve(dbPath);
    this.memoryStore = {
      conversations: new Map(),
      entries: new Map(),
      tasks: new Map(),
      documents: new Map(),
    };
  }

  async open(): Promise<void> {
    const dir = path.dirname(this.dbPath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    // Try opening node:sqlite if available (Node 22+)
    try {
      // Dynamic import to support runtimes gracefully
      const { DatabaseSync } = await import('node:sqlite');
      this.db = new DatabaseSync(this.dbPath);
      this.initSchema();
      return;
    } catch {
      // Fallback to JSON-backed durable store if node:sqlite is unavailable
      this.loadJsonStore();
    }
  }

  private initSchema(): void {
    if (!this.db) return;

    this.db.exec(`
      CREATE TABLE IF NOT EXISTS conversations (
        id TEXT PRIMARY KEY,
        parent_id TEXT,
        fork_entry_id TEXT,
        title TEXT NOT NULL,
        model TEXT NOT NULL,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS entries (
        id TEXT PRIMARY KEY,
        conversation_id TEXT NOT NULL,
        role TEXT NOT NULL,
        content TEXT,
        tool_calls TEXT,
        tool_call_id TEXT,
        created_at TEXT NOT NULL
      );
      CREATE INDEX IF NOT EXISTS idx_entries_conv ON entries(conversation_id);

      CREATE TABLE IF NOT EXISTS tasks (
        id TEXT PRIMARY KEY,
        conversation_id TEXT NOT NULL,
        name TEXT NOT NULL,
        status TEXT NOT NULL,
        checkpoint TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS documents (
        kind TEXT NOT NULL,
        scope_id TEXT NOT NULL,
        version INTEGER NOT NULL,
        data TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        PRIMARY KEY (kind, scope_id)
      );
    `);
  }

  private loadJsonStore(): void {
    const jsonPath = this.dbPath.endsWith('.sqlite')
      ? this.dbPath.replace(/\.sqlite$/, '.json')
      : `${this.dbPath}.json`;

    if (fs.existsSync(jsonPath)) {
      try {
        const raw = fs.readFileSync(jsonPath, 'utf8');
        const parsed = JSON.parse(raw);
        if (parsed.conversations) {
          for (const c of parsed.conversations) this.memoryStore.conversations.set(c.id, c);
        }
        if (parsed.entries) {
          for (const [k, v] of Object.entries(parsed.entries)) {
            this.memoryStore.entries.set(k, v as StorageEntry[]);
          }
        }
        if (parsed.tasks) {
          for (const t of parsed.tasks) this.memoryStore.tasks.set(t.id, t);
        }
        if (parsed.documents) {
          for (const d of parsed.documents) {
            this.memoryStore.documents.set(`${d.kind}:${d.scopeId}`, d as StorageDocument);
          }
        }
      } catch {
        // Ignore read/parse error
      }
    }
  }

  private flushJsonStore(): void {
    if (this.db) return;
    const jsonPath = this.dbPath.endsWith('.sqlite')
      ? this.dbPath.replace(/\.sqlite$/, '.json')
      : `${this.dbPath}.json`;

    const data = {
      conversations: Array.from(this.memoryStore.conversations.values()),
      entries: Object.fromEntries(this.memoryStore.entries.entries()),
      tasks: Array.from(this.memoryStore.tasks.values()),
      documents: Array.from(this.memoryStore.documents.values()),
    };
    fs.writeFileSync(jsonPath, JSON.stringify(data, null, 2), 'utf8');
  }

  async saveConversation(conv: StorageConversation): Promise<void> {
    if (this.db) {
      const stmt = this.db.prepare(`
        INSERT INTO conversations (id, parent_id, fork_entry_id, title, model, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
          title = excluded.title,
          model = excluded.model,
          updated_at = excluded.updated_at
      `);
      stmt.run(conv.id, conv.parentId || null, conv.forkEntryId || null, conv.title, conv.model, conv.createdAt, conv.updatedAt);
    } else {
      this.memoryStore.conversations.set(conv.id, conv);
      this.flushJsonStore();
    }
  }

  async getConversation(id: string): Promise<StorageConversation | null> {
    if (this.db) {
      const stmt = this.db.prepare(`SELECT * FROM conversations WHERE id = ?`);
      const row = stmt.get(id) as any;
      if (!row) return null;
      return {
        id: row.id,
        parentId: row.parent_id || undefined,
        forkEntryId: row.fork_entry_id || undefined,
        title: row.title,
        model: row.model,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
      };
    }
    return this.memoryStore.conversations.get(id) || null;
  }

  async listConversations(): Promise<StorageConversation[]> {
    if (this.db) {
      const stmt = this.db.prepare(`SELECT * FROM conversations ORDER BY updated_at DESC`);
      const rows = stmt.all() as any[];
      return rows.map((row) => ({
        id: row.id,
        parentId: row.parent_id || undefined,
        forkEntryId: row.fork_entry_id || undefined,
        title: row.title,
        model: row.model,
        createdAt: row.created_at,
        updatedAt: row.updated_at,
      }));
    }
    return Array.from(this.memoryStore.conversations.values()).sort(
      (a, b) => b.updatedAt.localeCompare(a.updatedAt)
    );
  }

  async appendEntry(entry: StorageEntry): Promise<void> {
    if (this.db) {
      const stmt = this.db.prepare(`
        INSERT INTO entries (id, conversation_id, role, content, tool_calls, tool_call_id, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
      `);
      stmt.run(
        entry.id,
        entry.conversationId,
        entry.role,
        entry.content || null,
        entry.toolCalls ? JSON.stringify(entry.toolCalls) : null,
        entry.toolCallId || null,
        entry.createdAt
      );
    } else {
      const list = this.memoryStore.entries.get(entry.conversationId) || [];
      list.push(entry);
      this.memoryStore.entries.set(entry.conversationId, list);
      this.flushJsonStore();
    }
  }

  async getEntries(conversationId: string): Promise<StorageEntry[]> {
    if (this.db) {
      const stmt = this.db.prepare(`
        SELECT * FROM entries WHERE conversation_id = ? ORDER BY created_at ASC, rowid ASC
      `);
      const rows = stmt.all(conversationId) as any[];
      return rows.map((r) => ({
        id: r.id,
        conversationId: r.conversation_id,
        role: r.role,
        content: r.content,
        toolCalls: r.tool_calls ? JSON.parse(r.tool_calls) : undefined,
        toolCallId: r.tool_call_id || undefined,
        createdAt: r.created_at,
      }));
    }
    return this.memoryStore.entries.get(conversationId) || [];
  }

  async saveDocument<T>(doc: StorageDocument<T>): Promise<void> {
    if (this.db) {
      const stmt = this.db.prepare(`
        INSERT INTO documents (kind, scope_id, version, data, updated_at)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(kind, scope_id) DO UPDATE SET
          version = excluded.version,
          data = excluded.data,
          updated_at = excluded.updated_at
      `);
      stmt.run(doc.kind, doc.scopeId, doc.version, JSON.stringify(doc.data), doc.updatedAt);
    } else {
      this.memoryStore.documents.set(`${doc.kind}:${doc.scopeId}`, doc as StorageDocument);
      this.flushJsonStore();
    }
  }

  async getDocument<T>(kind: string, scopeId: string): Promise<StorageDocument<T> | null> {
    if (this.db) {
      const stmt = this.db.prepare(`SELECT * FROM documents WHERE kind = ? AND scope_id = ?`);
      const row = stmt.get(kind, scopeId) as any;
      if (!row) return null;
      return {
        kind: row.kind,
        scopeId: row.scope_id,
        version: row.version,
        data: JSON.parse(row.data) as T,
        updatedAt: row.updated_at,
      };
    }
    const doc = this.memoryStore.documents.get(`${kind}:${scopeId}`);
    return doc ? (doc as StorageDocument<T>) : null;
  }

  async listDocuments(scopeId?: string): Promise<StorageDocument[]> {
    if (this.db) {
      const stmt = scopeId
        ? this.db.prepare(`SELECT * FROM documents WHERE scope_id = ?`)
        : this.db.prepare(`SELECT * FROM documents`);
      const rows = (scopeId ? stmt.all(scopeId) : stmt.all()) as any[];
      return rows.map((r) => ({
        kind: r.kind,
        scopeId: r.scope_id,
        version: r.version,
        data: JSON.parse(r.data),
        updatedAt: r.updated_at,
      }));
    }
    const docs = Array.from(this.memoryStore.documents.values());
    return scopeId ? docs.filter((d) => d.scopeId === scopeId) : docs;
  }

  async saveTask(task: StorageTask): Promise<void> {
    if (this.db) {
      const stmt = this.db.prepare(`
        INSERT INTO tasks (id, conversation_id, name, status, checkpoint, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
          status = excluded.status,
          checkpoint = excluded.checkpoint,
          updated_at = excluded.updated_at
      `);
      stmt.run(
        task.id,
        task.conversationId,
        task.name,
        task.status,
        task.checkpoint || null,
        task.createdAt,
        task.updatedAt
      );
    } else {
      this.memoryStore.tasks.set(task.id, task);
      this.flushJsonStore();
    }
  }

  async getUnfinishedTasks(): Promise<StorageTask[]> {
    if (this.db) {
      const stmt = this.db.prepare(`
        SELECT * FROM tasks WHERE status IN ('pending', 'running')
      `);
      const rows = stmt.all() as any[];
      return rows.map((r) => ({
        id: r.id,
        conversationId: r.conversation_id,
        name: r.name,
        status: r.status,
        checkpoint: r.checkpoint || undefined,
        createdAt: r.created_at,
        updatedAt: r.updated_at,
      }));
    }
    return Array.from(this.memoryStore.tasks.values()).filter(
      (t) => t.status === 'pending' || t.status === 'running'
    );
  }

  close(): void {
    if (this.db) {
      try {
        this.db.close();
      } catch {
        // Ignore
      }
      this.db = null;
    }
  }
}
