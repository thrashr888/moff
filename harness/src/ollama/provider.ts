import type { CompletionRequest, CompletionResponse, StreamDelta } from '../copilot/provider.js';
import { randomUUID } from 'node:crypto';
import { z } from 'zod';
import { hasRepetitionLoop } from './repetition.js';

export const DEFAULT_OLLAMA_MODEL = 'qwen3.5:9b';

const toolCallSchema = z.object({
  id: z.string(), type: z.literal('function'),
  function: z.object({ name: z.string(), arguments: z.string() }),
});
const completionSchema = z.object({
  id: z.string(), model: z.string(),
  choices: z.array(z.object({
    message: z.object({
      role: z.literal('assistant'), content: z.string().nullable(),
      tool_calls: z.array(toolCallSchema).optional(),
    }),
    finish_reason: z.string(),
  })).min(1),
  usage: z.object({
    prompt_tokens: z.number(), completion_tokens: z.number(), total_tokens: z.number(),
  }).optional(),
});
const chunkSchema = z.object({
  id: z.string(),
  choices: z.array(z.object({
    finish_reason: z.string().nullable().optional(),
    delta: z.object({
      content: z.string().nullable().optional(),
      tool_calls: z.array(z.object({
        index: z.number().int().nonnegative(), id: z.string().optional(),
        type: z.literal('function').optional(),
        function: z.object({ name: z.string().optional(), arguments: z.string().optional() }).optional(),
      })).optional(),
    }),
  })),
});

export class OllamaProvider {
  private baseUrl: string;

  constructor(baseUrl: string = process.env.OLLAMA_HOST || 'http://127.0.0.1:11434') {
    this.baseUrl = baseUrl.replace(/\/+$/, '');
  }

  get isLocal(): boolean {
    return ['localhost', '127.0.0.1', '[::1]'].includes(new URL(this.baseUrl).hostname);
  }

  async listModels(): Promise<string[]> {
    try {
      const res = await fetch(`${this.baseUrl}/api/tags`);
      if (!res.ok) return ['llama3.2:latest', 'qwen3:latest'];
      const data = (await res.json()) as { models: Array<{ name: string; model: string }> };
      return data.models.map((m) => m.name || m.model);
    } catch {
      return ['llama3.2:latest', 'qwen3:latest'];
    }
  }

  async complete(request: CompletionRequest): Promise<CompletionResponse> {
    const model = request.model || DEFAULT_OLLAMA_MODEL;
    const url = `${this.baseUrl}/v1/chat/completions`;

    const body: Record<string, unknown> = {
      model,
      messages: request.messages.map(message => ({ ...message, content: message.content ?? '' })),
      temperature: request.temperature ?? 0.2,
      stream: false,
      ...(/^qwen3(?:\.\d+)?(?::|$)/.test(model) ? { reasoning_effort: 'none' } : {}),
    };

    if (request.tools && request.tools.length > 0) {
      body.tools = request.tools;
      body.tool_choice = 'auto';
    }

    if (request.maxTokens !== undefined) body.max_tokens = request.maxTokens;

    const res = await fetch(url, {
      method: 'POST',
      signal: request.signal,
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`Ollama request failed (${res.status}): ${errText}`);
    }

    const data = completionSchema.parse(await res.json());
    const choice = data.choices[0];
    if (!choice.message.content?.trim() && !choice.message.tool_calls?.length) {
      throw new Error(`Ollama model ${model} returned neither answer text nor tool calls. The server may have discarded malformed tool arguments, or the model may have returned only reasoning. No action was executed by this response.`);
    }
    return {
      id: data.id || randomUUID(),
      model: data.model || model,
      message: choice.message,
      usage: data.usage,
      finish_reason: choice?.finish_reason,
    };
  }

  async *stream(request: CompletionRequest): AsyncGenerator<StreamDelta, CompletionResponse, void> {
    const model = request.model || DEFAULT_OLLAMA_MODEL;
    const url = `${this.baseUrl}/v1/chat/completions`;

    const body: Record<string, unknown> = {
      model,
      messages: request.messages.map(message => ({ ...message, content: message.content ?? '' })),
      temperature: request.temperature ?? 0.2,
      stream: true,
      ...(/^qwen3(?:\.\d+)?(?::|$)/.test(model) ? { reasoning_effort: 'none' } : {}),
    };

    if (request.tools && request.tools.length > 0) {
      body.tools = request.tools;
      body.tool_choice = 'auto';
    }

    if (request.maxTokens !== undefined) body.max_tokens = request.maxTokens;

    const res = await fetch(url, {
      method: 'POST',
      signal: request.signal,
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const errText = await res.text();
      throw new Error(`Ollama streaming request failed (${res.status}): ${errText}`);
    }

    if (!res.body) {
      throw new Error('Ollama streaming response body is empty');
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    let fullContent = '';
    let responseId: string = randomUUID();
    let finishReason: string | undefined;
    const toolCalls = new Map<number, { id: string; name: string; arguments: string }>();
    const sourceText = request.messages
      .filter(message => message.role === 'tool')
      .map(message => message.content || '').join('\n');

    try {
      let ended = false;
      while (!ended) {
        const { done, value } = await reader.read();
        buffer += done ? decoder.decode() : decoder.decode(value, { stream: true });
        if (done && buffer) buffer += '\n';
        const lines = buffer.split('\n');
        buffer = lines.pop() || '';

        for (const line of lines) {
          const trimmed = line.trim();
          if (!trimmed || trimmed.startsWith(':')) continue;
          if (trimmed === 'data: [DONE]') { ended = true; break; }
          if (trimmed.startsWith('data: ')) {
            const json = chunkSchema.parse(JSON.parse(trimmed.slice(6)));
            responseId = json.id || responseId;
            const choice = json.choices?.[0];
            if (!choice) continue;
            if (choice.finish_reason) finishReason = choice.finish_reason;
            const delta = choice.delta;
            for (const call of delta.tool_calls || []) {
              const accumulated = toolCalls.get(call.index) || { id: '', name: '', arguments: '' };
              if (call.id) accumulated.id = call.id;
              accumulated.name += call.function?.name || '';
              accumulated.arguments += call.function?.arguments || '';
              toolCalls.set(call.index, accumulated);
            }
            if (delta?.content) {
              fullContent += delta.content;
              if (hasRepetitionLoop(fullContent, sourceText)) {
                throw new Error('Stopped repetitive model output, including repeated entries with unsupported timestamps. The partial answer is incomplete and may contain invented records.');
              }
            }
            yield {
              ...(delta.content ? { content: delta.content } : {}),
              ...(delta.tool_calls ? { tool_calls: delta.tool_calls } : {}),
            };
          }
        }
        if (done) break;
      }
      if (!finishReason) throw new Error('Ollama stream ended before a completion finish reason');
      const calls = [...toolCalls.entries()].sort(([a], [b]) => a - b).map(([, call]) => {
        if (!call.id || !call.name) throw new Error('Incomplete Ollama tool call');
        return { id: call.id, type: 'function' as const, function: { name: call.name, arguments: call.arguments } };
      });
      if (!fullContent.trim() && !calls.length) {
        throw new Error(`Ollama model ${model} returned neither answer text nor tool calls. The server may have discarded malformed tool arguments, or the model may have returned only reasoning. No action was executed by this response.`);
      }
      return {
      id: responseId,
      model,
      message: {
        role: 'assistant',
        content: fullContent,
        ...(calls.length ? { tool_calls: calls } : {}),
      },
      finish_reason: finishReason,
      };
    } finally {
      await reader.cancel();
      reader.releaseLock();
    }
  }
}
