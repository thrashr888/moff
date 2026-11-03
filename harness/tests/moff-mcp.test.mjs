import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MoffMcp } from '../dist/tools/moff-mcp.js';

test('SDK discovers native Moff tool and CLI defaults; writes are rejected', async () => {
  const moff = new MoffMcp();
  try {
    await moff.connect();
    assert.deepEqual(moff.tools.map(tool => tool.function.name), ['moff']);
    assert.match(moff.instructions, /moff.schema\/v1/);
    assert.match(await moff.call('moff', '{"args":["outlook","mail","list","--help"]}'), /default: 25/);
    await assert.rejects(moff.call('moff', '{"args":["auth","refresh","--confirm"]}'), /read-only/);
    await assert.rejects(moff.call('moff_mail_list', '{}'), /Unknown MCP tool/);
  } finally {
    await moff.close();
  }
});
