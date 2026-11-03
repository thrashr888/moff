import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

test('npm run chat uses Ollama despite BOT_PROVIDER=copilot and streams a reply', { timeout: 120_000 }, async () => {
  const directory = await mkdtemp(join(tmpdir(), 'moff-cli-'));
  const child = spawn('npm', ['run', 'chat'], {
    env: {
      ...process.env,
      BOT_PROVIDER: 'copilot',
      BOT_MODEL: '',
      BOT_STORAGE_PATH: join(directory, 'bot.sqlite'),
    },
    stdio: ['pipe', 'pipe', 'pipe'],
  });
  let stdout = '';
  let stderr = '';
  let sent = false;
  let exiting = false;
  const timer = setTimeout(() => child.kill(), 110_000);
  child.stdout.on('data', chunk => {
    stdout += chunk.toString();
    if (!sent && stdout.includes("Type message or 'exit'.") && stdout.includes('> ')) {
      sent = true;
      child.stdin.write('Reply with exactly OLLAMA_CLI_OK and nothing else.\n');
    }
    if (sent && !exiting && stdout.includes('OLLAMA_CLI_OK') && stdout.endsWith('> ')) {
      exiting = true;
      child.stdin.end('exit\n');
    }
  });
  child.stderr.on('data', chunk => { stderr += chunk.toString(); });
  try {
    const code = await new Promise((resolve, reject) => {
      child.on('error', reject);
      child.on('close', resolve);
    });
    assert.equal(code, 0, `${stdout}\n${stderr}`);
    assert.match(stdout, /provider: ollama, model: qwen3\.5:9b/);
    assert.match(stdout, /OLLAMA_CLI_OK/);
    assert.doesNotMatch(`${stdout}\n${stderr}`, /403|Error:|Copilot token/);
  } finally {
    clearTimeout(timer);
    if (child.exitCode === null) child.kill();
    await rm(directory, { recursive: true, force: true });
  }
});
