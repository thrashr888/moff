export interface CopilotModelInfo {
  id: string;
  name: string;
  provider: 'copilot';
  contextWindow: number;
  maxOutputTokens: number;
  supportsTools: boolean;
  supportsStreaming: boolean;
  supportsThinking: boolean;
  defaultThinkingBudget?: number;
}

export const COPILOT_MODELS: Record<string, CopilotModelInfo> = {
  'claude-3.5-sonnet': {
    id: 'claude-3.5-sonnet',
    name: 'Claude 3.5 Sonnet (Copilot)',
    provider: 'copilot',
    contextWindow: 200000,
    maxOutputTokens: 8192,
    supportsTools: true,
    supportsStreaming: true,
    supportsThinking: false,
  },
  'claude-3.7-sonnet': {
    id: 'claude-3.7-sonnet',
    name: 'Claude 3.7 Sonnet (Copilot)',
    provider: 'copilot',
    contextWindow: 200000,
    maxOutputTokens: 16384,
    supportsTools: true,
    supportsStreaming: true,
    supportsThinking: true,
    defaultThinkingBudget: 2048,
  },
  'gpt-4o': {
    id: 'gpt-4o',
    name: 'GPT-4o (Copilot)',
    provider: 'copilot',
    contextWindow: 128000,
    maxOutputTokens: 4096,
    supportsTools: true,
    supportsStreaming: true,
    supportsThinking: false,
  },
  'gpt-4o-mini': {
    id: 'gpt-4o-mini',
    name: 'GPT-4o mini (Copilot)',
    provider: 'copilot',
    contextWindow: 128000,
    maxOutputTokens: 4096,
    supportsTools: true,
    supportsStreaming: true,
    supportsThinking: false,
  },
  'o1': {
    id: 'o1',
    name: 'o1 Reasoning (Copilot)',
    provider: 'copilot',
    contextWindow: 200000,
    maxOutputTokens: 16384,
    supportsTools: true,
    supportsStreaming: true,
    supportsThinking: true,
  },
  'o3-mini': {
    id: 'o3-mini',
    name: 'o3-mini (Copilot)',
    provider: 'copilot',
    contextWindow: 200000,
    maxOutputTokens: 16384,
    supportsTools: true,
    supportsStreaming: true,
    supportsThinking: true,
  },
};

export const DEFAULT_COPILOT_MODEL = 'claude-3.5-sonnet';
