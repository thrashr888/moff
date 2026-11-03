import { describe, it } from 'node:test';
import * as assert from 'node:assert';
import { CopilotAuthManager } from '../src/copilot/auth.js';
import { CopilotProvider } from '../src/copilot/provider.js';
import { COPILOT_MODELS, DEFAULT_COPILOT_MODEL } from '../src/copilot/models.js';

describe('Copilot Provider & Models', () => {
  it('should expose defined Copilot models with correct properties', () => {
    assert.ok(COPILOT_MODELS['claude-3.5-sonnet']);
    assert.strictEqual(COPILOT_MODELS['claude-3.5-sonnet'].contextWindow, 200000);
    assert.strictEqual(COPILOT_MODELS['claude-3.5-sonnet'].supportsTools, true);
    assert.strictEqual(DEFAULT_COPILOT_MODEL, 'claude-3.5-sonnet');
    assert.ok(COPILOT_MODELS['gpt-4o']);
    assert.ok(COPILOT_MODELS['o3-mini']);
  });

  it('should initialize CopilotAuthManager with explicit token', async () => {
    const auth = new CopilotAuthManager('test_token');
    assert.ok(auth);
  });

  it('should adapt to Pi-AI provider format', () => {
    const provider = new CopilotProvider();
    const piProvider = provider.asPiAiProvider();
    assert.strictEqual(piProvider.name, 'copilot');
    assert.ok(piProvider.models.length >= 4);
    const claudeModel = piProvider.models.find((m) => m.id === 'claude-3.5-sonnet');
    assert.ok(claudeModel);
    assert.strictEqual(claudeModel.contextWindow, 200000);
  });
});
