import { callAI, type ChatMessage, type ToolCall, type ToolResult, type AIModelConfig } from './ai-providers';
import { excelTools, executeToolCall } from './agent-tools';

export interface PlanStep {
  id: number;
  description: string;
  tool: string;
  toolArgs: Record<string, any>;
}

export interface Plan {
  steps: PlanStep[];
  reasoning: string;
}

export interface ExecutionResult {
  step: PlanStep;
  success: boolean;
  output: string;
  error?: string;
}

export interface AgentLog {
  agent: 'planner' | 'executor' | 'validator';
  content: string;
  timestamp: number;
}

const PLANNER_PROMPT = `You are the PLANNER agent for an Excel AI assistant. Your job is to break down user requests into a clear, step-by-step execution plan.

## Rules:
1. ALWAYS start by analyzing what information you need. If you don't know the sheet structure or current data, plan to read it first.
2. Each step should use EXACTLY ONE tool call.
3. Steps must be sequential and build on previous results.
4. For data operations: READ first, then WRITE/MODIFY.
5. For formatting: READ the range first to understand current state, then apply format.
6. For complex tasks (create table + format + fill data): plan ALL steps upfront.
7. Use specific cell references (A1 notation).
8. If the user mentions "selected range", plan to call get_selected_range first.

## Available Tools:
- get_workbook_structure: Get all sheets, tables, named ranges
- get_selected_range: Get data from current selection
- get_range: Get data from a specific range (e.g., "A1:C10")
- get_sheet_data: Get all data from a sheet
- set_values: Write values to a range
- set_formulas: Write formulas to a range
- apply_format: Apply formatting (bold, colors, alignment, number format)
- insert_rows / delete_rows / insert_columns / delete_columns
- add_worksheet / delete_worksheet
- create_table: Convert a range to an Excel Table
- sort_range: Sort a range by column
- auto_fill: Autofill a formula/pattern to a target range

## Output Format:
Respond with a JSON object ONLY:
{
  "reasoning": "Brief explanation of your approach",
  "steps": [
    {
      "id": 1,
      "description": "What this step does",
      "tool": "tool_name",
      "toolArgs": { "arg1": "value1" }
    }
  ]
}

If the task is simple (just reading data), you can have 1-2 steps.
If the task is complex, plan ALL steps needed.`;

const VALIDATOR_PROMPT = `You are the VALIDATOR agent for an Excel AI assistant. Your job is to review execution results and determine if the task is complete or needs retry.

## Rules:
1. Check if ALL planned steps were executed successfully.
2. If any step failed, determine if it's critical or can be skipped.
3. If the task is complete, respond with a natural summary of what was done.
4. If the task is NOT complete, create a RETRY plan with specific steps.
5. Be lenient — if partial success achieves the user's goal, mark as complete.
6. If a tool returned an error, try to understand WHY and plan an alternative approach.

## Output Format:
If task is COMPLETE, respond with JSON:
{
  "status": "complete",
  "summary": "Natural language summary of what was accomplished"
}

If task needs RETRY, respond with JSON:
{
  "status": "retry",
  "reasoning": "Why the task is not complete",
  "steps": [
    {
      "id": 1,
      "description": "What this retry step does",
      "tool": "tool_name",
      "toolArgs": { "arg1": "value1" }
    }
  ]
}`;

function buildToolCallsArg(steps: PlanStep[]): string {
  return JSON.stringify(steps.map((s) => ({ tool: s.tool, args: s.toolArgs })));
}

export async function runPlanner(
  userMessage: string,
  config: AIModelConfig,
  conversationContext: string
): Promise<Plan> {
  const messages: ChatMessage[] = [
    { role: 'system', content: PLANNER_PROMPT },
    { role: 'user', content: `User request: ${userMessage}\n\nConversation context: ${conversationContext || 'No previous context.'}` },
  ];

  const response = await callAI(messages, config);

  try {
    const jsonStr = response.content.match(/\{[\s\S]*\}/)?.[0] || response.content;
    const plan = JSON.parse(jsonStr) as Plan;
    return plan;
  } catch {
    return {
      steps: [
        {
          id: 1,
          description: 'Get workbook structure to understand the current state',
          tool: 'get_workbook_structure',
          toolArgs: {},
        },
      ],
      reasoning: 'Fallback: getting workbook structure first',
    };
  }
}

export async function runExecutor(
  plan: Plan,
  config: AIModelConfig,
  onStepComplete?: (step: PlanStep, result: ExecutionResult) => void
): Promise<ExecutionResult[]> {
  const results: ExecutionResult[] = [];

  for (const step of plan.steps) {
    try {
      const output = await executeToolCall({
        id: `step-${step.id}`,
        name: step.tool,
        arguments: JSON.stringify(step.toolArgs),
      });

      const result: ExecutionResult = {
        step,
        success: !output.startsWith('Error'),
        output,
        error: output.startsWith('Error') ? output : undefined,
      };

      results.push(result);
      onStepComplete?.(step, result);
    } catch (error) {
      const result: ExecutionResult = {
        step,
        success: false,
        output: '',
        error: error instanceof Error ? error.message : String(error),
      };

      results.push(result);
      onStepComplete?.(step, result);
    }
  }

  return results;
}

export async function runValidator(
  userMessage: string,
  plan: Plan,
  results: ExecutionResult[],
  config: AIModelConfig
): Promise<{ status: 'complete' | 'retry'; summary?: string; reasoning?: string; steps?: PlanStep[] }> {
  const resultsSummary = results
    .map(
      (r) =>
        `Step ${r.step.id} (${r.step.tool}): ${r.success ? 'SUCCESS' : 'FAILED'}\n  Output: ${r.output.slice(0, 300)}`
    )
    .join('\n\n');

  const messages: ChatMessage[] = [
    { role: 'system', content: VALIDATOR_PROMPT },
    {
      role: 'user',
      content: `Original request: ${userMessage}\n\nPlan: ${JSON.stringify(plan, null, 2)}\n\nExecution results:\n${resultsSummary}`,
    },
  ];

  const response = await callAI(messages, config);

  try {
    const jsonStr = response.content.match(/\{[\s\S]*\}/)?.[0] || response.content;
    return JSON.parse(jsonStr);
  } catch {
    return {
      status: 'complete',
      summary: response.content || 'Task completed.',
    };
  }
}

export async function runResponder(
  userMessage: string,
  plan: Plan,
  results: ExecutionResult[],
  config: AIModelConfig
): Promise<string> {
  const resultsSummary = results
    .map(
      (r) =>
        `Step ${r.step.id} (${r.step.description}): ${r.success ? 'Done' : 'Failed - ' + r.error}`
    )
    .join('\n');

  const messages: ChatMessage[] = [
    {
      role: 'system',
      content: `You are the final response generator for an Excel AI assistant. Summarize what was accomplished for the user in a clear, friendly way.

## Rules:
1. Be concise but informative.
2. Mention specific cell ranges, sheet names, and what was changed.
3. If something failed, explain what went wrong and suggest alternatives.
4. Use natural language, not technical jargon.
5. If formulas were created, explain what they do.`,
    },
    {
      role: 'user',
      content: `User asked: ${userMessage}\n\nPlan: ${plan.reasoning}\n\nResults:\n${resultsSummary}`,
    },
  ];

  const response = await callAI(messages, config);
  return response.content;
}
