import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';

test('ACP EOF waits for requests and closes MCP subprocess', { timeout: 15_000 }, async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-eof-'));
  const child = spawn(process.execPath, ['dist/cli.js', 'acp'], {
    env: { ...process.env, BOT_PROVIDER: 'ollama', BOT_STORAGE_PATH: join(directory, 'bot.sqlite') },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  let stdout = '';
  let stderr = '';
  child.stdout.on('data', chunk => { stdout += chunk; });
  child.stderr.on('data', chunk => { stderr += chunk; });
  const timer = setTimeout(() => child.kill(), 10_000);
  try {
    child.stdin.end('{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":1,"clientCapabilities":{}}}\n');
    const code = await new Promise((resolve, reject) => {
      child.on('error', reject);
      child.on('close', resolve);
    });
    assert.equal(code, 0, stderr);
    assert.equal(JSON.parse(stdout.trim()).id, 1);
    assert.equal(JSON.parse(stdout.trim()).result.protocolVersion, 1);
  } finally {
    clearTimeout(timer);
    await rm(directory, { recursive: true, force: true });
  }
});
