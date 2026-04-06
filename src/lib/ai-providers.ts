export type AIProvider = 'openai' | 'anthropic' | 'google' | 'openrouter';

export interface AIModelConfig {
  provider: AIProvider;
  model: string;
  apiKey: string;
  endpoint?: string;
}

export interface ChatMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
}

export interface ToolResult {
  tool_call_id: string;
  content: string;
}

export interface AIResponse {
  content: string;
  toolCalls?: ToolCall[];
  usage?: {
    promptTokens: number;
    completionTokens: number;
    totalTokens: number;
  };
}

export async function callAI(
  messages: ChatMessage[],
  config: AIModelConfig,
  tools?: any[]
): Promise<AIResponse> {
  switch (config.provider) {
    case 'openai':
      return callOpenAI(messages, config, tools);
    case 'anthropic':
      return callAnthropic(messages, config, tools);
    case 'google':
      return callGoogle(messages, config, tools);
    case 'openrouter':
      return callOpenRouter(messages, config, tools);
    default:
      throw new Error(`Unknown provider: ${config.provider}`);
  }
}

async function callOpenAI(
  messages: ChatMessage[],
  config: AIModelConfig,
  tools?: any[]
): Promise<AIResponse> {
  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
    },
    body: JSON.stringify({
      model: config.model,
      messages,
      tools: tools?.length ? tools : undefined,
      tool_choice: tools?.length ? 'auto' : undefined,
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`OpenAI API error: ${response.status} - ${text}`);
  }

  const data = await response.json();
  const choice = data.choices[0].message;

  const toolCalls: ToolCall[] | undefined = choice.tool_calls?.map((tc: any) => ({
    id: tc.id,
    name: tc.function.name,
    arguments: tc.function.arguments,
  }));

  return {
    content: choice.content || '',
    toolCalls,
    usage: data.usage,
  };
}

async function callOpenRouter(
  messages: ChatMessage[],
  config: AIModelConfig,
  tools?: any[]
): Promise<AIResponse> {
  const endpoint = config.endpoint || 'https://openrouter.ai/api/v1/chat/completions';

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${config.apiKey}`,
      'HTTP-Referer': 'https://localhost:3000',
      'X-Title': 'Cel',
    },
    body: JSON.stringify({
      model: config.model,
      messages,
      tools: tools?.length ? tools : undefined,
      tool_choice: tools?.length ? 'auto' : undefined,
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`OpenRouter API error: ${response.status} - ${text}`);
  }

  const data = await response.json();
  const choice = data.choices?.[0]?.message;

  if (!choice) {
    throw new Error(`OpenRouter returned no choices: ${JSON.stringify(data)}`);
  }

  const toolCalls: ToolCall[] | undefined = choice.tool_calls?.map((tc: any) => ({
    id: tc.id,
    name: tc.function.name,
    arguments: tc.function.arguments,
  }));

  return {
    content: choice.content || '',
    toolCalls,
    usage: data.usage,
  };
}

async function callAnthropic(
  messages: ChatMessage[],
  config: AIModelConfig,
  tools?: any[]
): Promise<AIResponse> {
  const systemMessage = messages.find((m) => m.role === 'system');
  const userMessages = messages.filter((m) => m.role !== 'system');

  const anthropicTools = tools?.map((t) => ({
    name: t.function.name,
    description: t.function.description,
    input_schema: t.function.parameters,
  }));

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': config.apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: config.model,
      max_tokens: 4096,
      system: systemMessage?.content,
      messages: userMessages.map((m) => ({
        role: m.role === 'assistant' ? 'assistant' : 'user',
        content: m.content,
      })),
      tools: anthropicTools,
    }),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Anthropic API error: ${response.status} - ${text}`);
  }

  const data = await response.json();
  let content = '';
  const toolCalls: ToolCall[] = [];

  for (const block of data.content) {
    if (block.type === 'text') {
      content += block.text;
    } else if (block.type === 'tool_use') {
      toolCalls.push({
        id: block.id,
        name: block.name,
        arguments: JSON.stringify(block.input),
      });
    }
  }

  return {
    content,
    toolCalls: toolCalls.length ? toolCalls : undefined,
    usage: data.usage
      ? {
          promptTokens: data.usage.input_tokens,
          completionTokens: data.usage.output_tokens,
          totalTokens: data.usage.input_tokens + data.usage.output_tokens,
        }
      : undefined,
  };
}

async function callGoogle(
  messages: ChatMessage[],
  config: AIModelConfig,
  tools?: any[]
): Promise<AIResponse> {
  const contents = messages
    .filter((m) => m.role !== 'system')
    .map((m) => ({
      role: m.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: m.content }],
    }));

  const systemMessage = messages.find((m) => m.role === 'system');

  const body: any = {
    contents,
    systemInstruction: systemMessage ? { parts: [{ text: systemMessage.content }] } : undefined,
  };

  if (tools?.length) {
    body.tools = tools.map((t) => ({
      functionDeclarations: [
        {
          name: t.function.name,
          description: t.function.description,
          parameters: t.function.parameters,
        },
      ],
    }));
  }

  const modelName = config.model.replace('gemini-', '');
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${config.apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    }
  );

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Google AI error: ${response.status} - ${text}`);
  }

  const data = await response.json();
  const candidate = data.candidates?.[0];
  const parts = candidate?.content?.parts || [];

  let content = '';
  const toolCalls: ToolCall[] = [];

  for (const part of parts) {
    if (part.text) {
      content += part.text;
    }
    if (part.functionCall) {
      toolCalls.push({
        id: `call-${Date.now()}`,
        name: part.functionCall.name,
        arguments: JSON.stringify(part.functionCall.args),
      });
    }
  }

  return {
    content,
    toolCalls: toolCalls.length ? toolCalls : undefined,
  };
}
