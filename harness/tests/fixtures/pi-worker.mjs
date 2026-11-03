import { PersistentBotHarness } from '../../dist/durable/harness.js';
import { readFile, writeFile } from 'node:fs/promises';

const path = process.env.PI_TEST_PATH;
const interrupted = process.env.PI_TEST_MODE === 'interrupt';
const provider = {
  async complete() { throw new Error('Unexpected nonstreaming call'); },
  async *stream(request) {
    if (!request.messages.some(message => message.role === 'tool')) {
      return {
        id: 'call', model: 'test', finish_reason: 'tool_calls',
        message: { role: 'assistant', content: '', tool_calls: [{
          id: 'schema', type: 'function', function: { name: 'moff', arguments: '{"args":["schema"]}' },
        }] },
      };
    }
    if (interrupted) {
      process.send({ type: 'pending' });
      await new Promise((_, reject) => {
        const timer = setInterval(() => {}, 1000);
        request.signal.addEventListener('abort', () => {
          clearInterval(timer);
          reject(new Error('Interrupted'));
        }, { once: true });
      });
    }
    yield { content: 'RECOVERED' };
    return { id: 'answer', model: 'test', finish_reason: 'stop',
      message: { role: 'assistant', content: 'RECOVERED' } };
  },
};

const harness = new PersistentBotHarness({ provider, defaultModel: 'test',
  storagePath: path, enableMoff: true });
try {
  await harness.init();
  const root = await harness.rootConversation();
  if (interrupted) {
    await writeFile(`${path}.root`, root.id);
    await harness.sendMessage(root.id, 'Describe the Moff schema.');
  } else {
    await harness.waitForIdle();
    const entries = await harness.getEntries(root.id);
    process.send({
      type: 'recovered', sameRoot: root.id === await readFile(`${path}.root`, 'utf8'),
      users: entries.filter(entry => entry.role === 'user').length,
      tools: entries.filter(entry => entry.role === 'tool').length,
      answer: entries.filter(entry => entry.role === 'assistant').at(-1).content,
    });
  }
} finally {
  await harness.close();
  process.disconnect();
}
