import { CopilotAuthManager } from './auth.js';
import { COPILOT_MODELS, DEFAULT_COPILOT_MODEL, CopilotModelInfo } from './models.js';
import { randomUUID } from 'node:crypto';

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant' | 'tool';
  content: string | null;
  name?: string;
  tool_call_id?: string;
  tool_calls?: Array<{
    id: string;
    type: 'function';
    function: {
      name: string;
      arguments: string;
    };
  }>;
}

export interface ToolDefinition {
  type: 'function';
  function: {
    name: string;
    description?: string;
    parameters: Record<string, unknown>;
  };
}

export interface CompletionRequest {
  model?: string;
  messages: ChatMessage[];
  tools?: ToolDefinition[];
  temperature?: number;
  maxTokens?: number;
  stream?: boolean;
  signal?: AbortSignal;
}

export interface CompletionResponse {
  id: string;
  model: string;
  message: ChatMessage;
  usage?: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
  finish_reason?: string;
}

export interface StreamDelta {
  content?: string;
  role?: string;
  tool_calls?: Array<{
    index: number;
    id?: string;
    type?: 'function';
    function?: {
      name?: string;
      arguments?: string;
    };
  }>;
}

export class CopilotProvider {
  private auth: CopilotAuthManager;

  constructor(auth?: CopilotAuthManager) {
    this.auth = auth || new CopilotAuthManager();
  }

  getModels(): Record<string, CopilotModelInfo> {
    return COPILOT_MODELS;
  }

  /**
   * Complete a chat conversation (non-streaming).
   */
  async complete(request: CompletionRequest): Promise<CompletionResponse> {
    const tokenInfo = await this.auth.getCopilotToken();
    const model = request.model || DEFAULT_COPILOT_MODEL;
    const url = `${tokenInfo.apiEndpoint}/chat/completions`;

    const body: Record<string, unknown> = {
      model,
      messages: request.messages,
      temperature: request.temperature ?? 0.2,
      stream: false,
    };

    if (request.tools && request.tools.length > 0) {
      body.tools = request.tools;
      body.tool_choice = 'auto';
    }

    if (request.maxTokens) {
      body.max_tokens = request.maxTokens;
    }

    const response = await fetch(url, {
      method: 'POST',
      signal: request.signal,
      headers: {
        Authorization: `Bearer ${tokenInfo.token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
        'X-Request-Id': randomUUID(),
        'Editor-Version': 'vscode/1.95.0',
        'Editor-Plugin-Version': 'copilot-chat/0.24.0',
        'Openai-Organization': 'github-copilot',
        'Copilot-Integration-Id': 'vscode-chat',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Copilot API request failed (${response.status}): ${errorText}`);
    }

    const data = (await response.json()) as {
      id: string;
      model: string;
      choices: Array<{
        message: ChatMessage;
        finish_reason: string;
      }>;
      usage?: {
        prompt_tokens: number;
        completion_tokens: number;
        total_tokens: number;
      };
    };

    const choice = data.choices[0];
    return {
      id: data.id,
      model: data.model,
      message: choice.message,
      usage: data.usage,
      finish_reason: choice.finish_reason,
    };
  }

  /**
   * Complete a chat conversation with streaming delta callbacks.
   */
  async *stream(request: CompletionRequest): AsyncGenerator<StreamDelta, CompletionResponse, void> {
    const tokenInfo = await this.auth.getCopilotToken();
    const model = request.model || DEFAULT_COPILOT_MODEL;
    const url = `${tokenInfo.apiEndpoint}/chat/completions`;

    const body: Record<string, unknown> = {
      model,
      messages: request.messages,
      temperature: request.temperature ?? 0.2,
      stream: true,
    };

    if (request.tools && request.tools.length > 0) {
      body.tools = request.tools;
      body.tool_choice = 'auto';
    }

    if (request.maxTokens) {
      body.max_tokens = request.maxTokens;
    }

    const response = await fetch(url, {
      method: 'POST',
      signal: request.signal,
      headers: {
        Authorization: `Bearer ${tokenInfo.token}`,
        'Content-Type': 'application/json',
        Accept: 'text/event-stream',
        'X-Request-Id': randomUUID(),
        'Editor-Version': 'vscode/1.95.0',
        'Editor-Plugin-Version': 'copilot-chat/0.24.0',
        'Openai-Organization': 'github-copilot',
        'Copilot-Integration-Id': 'vscode-chat',
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Copilot streaming request failed (${response.status}): ${errorText}`);
    }

    if (!response.body) {
      throw new Error('Copilot response body is empty');
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';

    let fullMessageContent = '';
    const fullToolCalls: Map<number, { id: string; name: string; arguments: string }> = new Map();
    let responseId = '';
    let responseModel = model;
    let finishReason = 'stop';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop() || '';

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith(':')) continue; // Skip empty and SSE comments

        if (trimmed === 'data: [DONE]') {
          break;
        }

        if (trimmed.startsWith('data: ')) {
          try {
            const json = JSON.parse(trimmed.slice(6));
            responseId = json.id || responseId;
            responseModel = json.model || responseModel;

            const choice = json.choices?.[0];
            if (!choice) continue;

            if (choice.finish_reason) {
              finishReason = choice.finish_reason;
            }

            const delta = choice.delta;
            if (delta) {
              if (delta.content) {
                fullMessageContent += delta.content;
              }
              if (delta.tool_calls) {
                for (const tc of delta.tool_calls) {
                  const idx = tc.index;
                  const existing = fullToolCalls.get(idx) || { id: '', name: '', arguments: '' };
                  if (tc.id) existing.id = tc.id;
                  if (tc.function?.name) existing.name += tc.function.name;
                  if (tc.function?.arguments) existing.arguments += tc.function.arguments;
                  fullToolCalls.set(idx, existing);
                }
              }
              yield delta;
            }
          } catch {
            // Partial SSE JSON, continue
          }
        }
      }
    }

    const toolCallsArray = Array.from(fullToolCalls.entries())
      .sort(([a], [b]) => a - b)
      .map(([, tc]) => ({
        id: tc.id,
        type: 'function' as const,
        function: {
          name: tc.name,
          arguments: tc.arguments,
        },
      }));

    return {
      id: responseId,
      model: responseModel,
      message: {
        role: 'assistant',
        content: fullMessageContent || null,
        ...(toolCallsArray.length > 0 ? { tool_calls: toolCallsArray } : {}),
      },
      finish_reason: finishReason,
    };
  }

  /**
   * Adapts this Copilot provider into a Pi-AI / Pi-Durable compatible provider object.
   */
  asPiAiProvider() {
    return {
      name: 'copilot',
      models: Object.values(COPILOT_MODELS).map((m) => ({
        id: m.id,
        name: m.name,
        contextWindow: m.contextWindow,
        maxOutputTokens: m.maxOutputTokens,
        reasoning: m.supportsThinking,
        tools: m.supportsTools,
      })),
      complete: async (req: any, _context?: any) => {
        // Convert pi-ai request to Copilot completion
        const messages: ChatMessage[] = (req.messages || []).map((msg: any) => ({
          role: msg.role,
          content: typeof msg.content === 'string' ? msg.content : JSON.stringify(msg.content),
          ...(msg.toolCallId ? { tool_call_id: msg.toolCallId } : {}),
          ...(msg.toolCalls ? { tool_calls: msg.toolCalls } : {}),
        }));

        const tools: ToolDefinition[] = (req.tools || []).map((t: any) => ({
          type: 'function',
          function: {
            name: t.name,
            description: t.description,
            parameters: t.parameters || {},
          },
        }));

        const res = await this.complete({
          model: req.model?.modelId || req.model || DEFAULT_COPILOT_MODEL,
          messages,
          tools: tools.length > 0 ? tools : undefined,
          temperature: req.temperature,
          maxTokens: req.maxTokens,
        });

        return {
          content: res.message.content
            ? [{ type: 'text', text: res.message.content }]
            : [],
          toolCalls: res.message.tool_calls?.map((tc) => ({
            id: tc.id,
            name: tc.function.name,
            arguments: JSON.parse(tc.function.arguments || '{}'),
          })),
          usage: res.usage,
        };
      },
    };
  }
}
