import { describe, it } from 'node:test';
import * as assert from 'node:assert';
import { PersistentBotHarness } from '../src/durable/harness.js';
import * as fs from 'node:fs';
import * as path from 'node:path';

describe('ACP Protocol Handling', () => {
  const testDb = path.resolve('./test-acp.sqlite');

  it('should initialize harness and create session for ACP', async () => {
    const harness = new PersistentBotHarness({ storagePath: testDb });
    await harness.init();

    const session = await harness.createConversation('ACP Session', 'claude-3.5-sonnet');
    assert.ok(session.id);
    assert.strictEqual(session.title, 'ACP Session');
    assert.strictEqual(session.model, 'claude-3.5-sonnet');

    harness.close();
    if (fs.existsSync(testDb)) fs.unlinkSync(testDb);
    if (fs.existsSync(`${testDb}.json`)) fs.unlinkSync(`${testDb}.json`);
    const jsonAlt = testDb.replace(/\.sqlite$/, '.json');
    if (fs.existsSync(jsonAlt)) fs.unlinkSync(jsonAlt);
  });
});
