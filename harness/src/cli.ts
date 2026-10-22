#!/usr/bin/env node
import { PersistentBotHarness } from './durable/harness.js';
import { DEFAULT_OLLAMA_MODEL, OllamaProvider } from './ollama/provider.js';
import { CopilotProvider } from './copilot/provider.js';
import * as readline from 'node:readline';
import { Readable, Writable } from 'node:stream';
import { ndJsonStream } from '@agentclientprotocol/sdk';
import { createAcpAgent } from './acp/agent.js';

async function main() {
  const args = process.argv.slice(2);
  let mode = 'chat';
  let model = process.env.BOT_MODEL || DEFAULT_OLLAMA_MODEL;
  let providerType = process.env.BOT_PROVIDER || 'ollama';

  for (let i = 0; i < args.length; i++) {
    if (args[i] === 'chat' || args[i] === 'acp' || args[i] === 'status') {
      mode = args[i];
    } else if (args[i] === '--model' && args[i + 1]) {
      model = args[++i];
    } else if (args[i] === '--copilot') {
      providerType = 'copilot';
    } else if (args[i] === '--ollama') {
      providerType = 'ollama';
    }
  }

  const provider = providerType === 'copilot' ? new CopilotProvider() : new OllamaProvider();
  const harness = new PersistentBotHarness({
    provider,
    providerId: providerType,
    defaultModel: model,
    storagePath: process.env.BOT_STORAGE_PATH,
  });
  await harness.init();

  if (mode === 'acp') {
    const connection = createAcpAgent(harness, process.cwd()).connect(
      ndJsonStream(Writable.toWeb(process.stdout), Readable.toWeb(process.stdin)));
    try { await connection.closed; }
    finally { await harness.close(); }
    return;
  }

  if (mode === 'status') {
    try { console.log(JSON.stringify(await harness.listConversations(), null, 2)); }
    finally { await harness.close(); }
    return;
  }

  console.log(`Persistent Bot (provider: ${providerType}, model: ${model}). Type message or 'exit'.`);
  const conv = await harness.rootConversation();
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  let closed = false;
  let busy = false;
  let closing: Promise<void> | undefined;
  const close = () => {
    closing ??= harness.close();
    void closing.catch(error => {
      console.error(`Shutdown failed: ${error instanceof Error ? error.message : String(error)}`);
      process.exitCode = 1;
    });
  };
  rl.on('close', () => {
    closed = true;
    if (!busy) close();
  });
  const promptUser = () => {
    if (closed) { close(); return; }
    rl.question('> ', async (input) => {
      const trimmed = input.trim();
      if (!trimmed || trimmed === 'exit') { rl.close(); close(); return; }
      busy = true;
      try {
        await harness.sendMessage(conv.id, trimmed, {
          model,
          onDelta: (d) => process.stdout.write(d),
          onTool: (message) => process.stderr.write(`\n${message}\n`),
        });
        console.log('\n');
      } catch (e: any) { console.error(`Error: ${e.message}`); }
      finally { busy = false; }
      promptUser();
    });
  };
  promptUser();
}
main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
