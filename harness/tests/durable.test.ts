import { describe, it } from 'node:test';
import * as assert from 'node:assert';
import { PersistentBotHarness } from '../src/durable/harness.js';
import { DurableStorage } from '../src/durable/storage.js';
import { TodoDoc, BotMemoryDoc } from '../src/durable/documents.js';
import * as fs from 'node:fs';
import * as path from 'node:path';

describe('Persistent Bot Harness & Storage', () => {
  const testDb = path.resolve('./test-storage.sqlite');

  it('should initialize durable storage and save documents', async () => {
    const storage = new DurableStorage(testDb);
    await storage.open();

    const doc = {
      kind: TodoDoc.kind,
      scopeId: 'test-conv',
      version: 1,
      data: {
        items: [{ id: '1', text: 'Test todo', completed: false, createdAt: new Date().toISOString() }],
      },
      updatedAt: new Date().toISOString(),
    };

    await storage.saveDocument(doc);
    const loaded = await storage.getDocument<typeof doc.data>(TodoDoc.kind, 'test-conv');
    assert.ok(loaded);
    assert.strictEqual(loaded.data.items.length, 1);
    assert.strictEqual(loaded.data.items[0].text, 'Test todo');

    storage.close();
  });

  it('should create and list conversations in harness', async () => {
    const harness = new PersistentBotHarness({ storagePath: testDb });
    await harness.init();

    const conv = await harness.createConversation('Persistent Bot Test');
    assert.ok(conv.id);
    assert.strictEqual(conv.title, 'Persistent Bot Test');

    const list = await harness.listConversations();
    assert.ok(list.length >= 1);
    const found = list.find((c) => c.id === conv.id);
    assert.ok(found);

    // Test fork
    const forked = await harness.forkConversation(conv.id);
    assert.strictEqual(forked.parentId, conv.id);

    harness.close();

    // Clean up test files
    if (fs.existsSync(testDb)) fs.unlinkSync(testDb);
    if (fs.existsSync(`${testDb}.json`)) fs.unlinkSync(`${testDb}.json`);
    const jsonAlt = testDb.replace(/\.sqlite$/, '.json');
    if (fs.existsSync(jsonAlt)) fs.unlinkSync(jsonAlt);
  });
});
