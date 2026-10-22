import {
  createAssistantMessageEventStream, createModels, createProvider,
  getCurrentSystemPrompt, getCurrentTools,
  type AssistantMessage, type Message, type Model, type ProviderStreams,
} from '@earendil-works/pi-ai';
import type { ChatMessage, CompletionRequest, CompletionResponse, StreamDelta } from '../copilot/provider.js';
import type { JsonObject } from '@earendil-works/pi-durable';

export interface ModelProvider {
  complete(request: CompletionRequest): Promise<CompletionResponse>;
  stream(request: CompletionRequest): AsyncGenerator<StreamDelta, CompletionResponse, void>;
}

function text(content: string | readonly { type: string; text?: string }[]): string {
  return typeof content === 'string' ? content : content.map(block => {
    if (block.type !== 'text') throw new Error(`The local model adapter does not support ${block.type} content`);
    return block.text || '';
  }).join('');
}

export function toChatMessage(message: Message): ChatMessage {
  if (message.role === 'toolResult') {
    return { role: 'tool',
      content: (message.isError ? 'Tool failed; this is not a successful data result:\n' : '') + text(message.content),
      tool_call_id: message.toolCallId };
  }
  if (message.role === 'assistant') {
    return {
      role: 'assistant',
      content: message.content.filter(block => block.type === 'text').map(block => block.text).join(''),
      tool_calls: message.content.filter(block => block.type === 'toolCall').map(call => ({
        id: call.id, type: 'function', function: { name: call.name, arguments: JSON.stringify(call.arguments) },
      })),
    };
  }
  return { role: message.role, content: text(message.content) };
}

// Adapt the tested transport, not the agent loop: Pi owns generation, tools, checkpoints and retries.
export function createBotModels(provider: ModelProvider, providerId: string, modelIds: string[]) {
  const models = createModels();
  const stream: ProviderStreams['stream'] = (model, context, options) => {
    const output = createAssistantMessageEventStream();
    const message: AssistantMessage = {
      role: 'assistant', content: [], api: model.api, provider: model.provider, model: model.id,
      timestamp: Date.now(), stopReason: 'stop',
      usage: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, totalTokens: 0,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 } },
    };
    void (async () => {
      try {
        const messages: ChatMessage[] = [
          { role: 'system', content: getCurrentSystemPrompt(context.messages) },
          ...context.messages.filter(item => item.role !== 'system').map(toChatMessage),
        ];
        const tools = getCurrentTools(context.messages).map(tool => ({
          type: 'function' as const,
          function: { name: tool.name, description: tool.description, parameters: { ...tool.parameters } },
        }));
        const source = provider.stream({
          model: model.id, messages, tools, signal: options?.signal,
          temperature: options?.temperature ?? 0.3,
        });
        output.push({ type: 'start', partial: message });
        let textStarted = false;
        while (true) {
          const next = await source.next();
          if (next.done) {
            const response = next.value;
            if (textStarted) {
              const block = message.content[0];
              if (block.type === 'text') output.push({ type: 'text_end', contentIndex: 0, content: block.text, partial: message });
            }
            for (const call of response.message.tool_calls || []) {
              const args: unknown = JSON.parse(call.function.arguments);
              if (!args || typeof args !== 'object' || Array.isArray(args)) throw new Error('Model tool arguments must be an object');
              const toolCall = { type: 'toolCall' as const, id: call.id, name: call.function.name,
                arguments: args as JsonObject };
              const contentIndex = message.content.length;
              message.content.push(toolCall);
              output.push({ type: 'toolcall_start', contentIndex, partial: message });
              output.push({ type: 'toolcall_end', contentIndex, toolCall, partial: message });
            }
            message.stopReason = response.finish_reason === 'length' ? 'length'
              : response.message.tool_calls?.length ? 'toolUse' : 'stop';
            if (response.usage) {
              message.usage = { ...message.usage, input: response.usage.prompt_tokens,
                output: response.usage.completion_tokens, totalTokens: response.usage.total_tokens };
            }
            output.push({ type: 'done', reason: message.stopReason, message });
            return;
          }
          if (next.value.content) {
            if (!textStarted) {
              message.content.push({ type: 'text', text: '' });
              output.push({ type: 'text_start', contentIndex: 0, partial: message });
              textStarted = true;
            }
            const block = message.content[0];
            if (block.type === 'text') block.text += next.value.content;
            output.push({ type: 'text_delta', contentIndex: 0, delta: next.value.content, partial: message });
          }
        }
      } catch (error) {
        message.stopReason = options?.signal?.aborted ? 'aborted' : 'error';
        message.errorMessage = error instanceof Error ? error.message : String(error);
        output.push({ type: 'error', reason: message.stopReason, error: message });
      }
    })();
    return output;
  };
  const catalog: Model<string>[] = modelIds.map(id => ({
    id, name: id, api: 'moff-chat', provider: providerId, baseUrl: '',
    reasoning: false, input: ['text'], contextWindow: 32768, maxTokens: 32768,
    cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 },
  }));
  models.setProvider(createProvider({
    id: providerId, models: catalog,
    auth: { apiKey: { name: providerId, resolve: async () => ({ auth: {} }) } },
    api: { stream, streamSimple: stream },
  }));
  return models;
}
