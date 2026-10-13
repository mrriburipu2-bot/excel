import Anthropic from "@anthropic-ai/sdk";
import { createClient, requestExtras } from "../shared/claude";
import { validate } from "../shared/schema";
import type { Settings } from "../shared/settings";
import { activeTools, type ToolContext, type ToolDef } from "./tools";

type MessageParam = Anthropic.Beta.BetaMessageParam;

export const SYSTEM_PROMPT = `You are an expert Excel assistant embedded in Microsoft Excel as an add-in. You act directly on the user's open workbook through tools.

You cover every area of Excel: formulas and functions (lookup, logical, text, date, statistical, financial, dynamic arrays such as FILTER, SORT, UNIQUE, XLOOKUP, LET, LAMBDA), data cleaning, tables, sorting and filtering, conditional formatting, data validation, charts, PivotTables, named ranges, formatting and layout, dashboards, financial models, what-if analysis, and explaining or debugging existing formulas.

How to work:
- Look before you act. If you don't know the layout, call get_workbook_info, then read the relevant ranges (or get_range_statistics for large data). Never guess addresses or header names.
- Prefer live formulas over hard-coded results so the workbook stays dynamic, unless the user asks for values.
- Use tools to make the changes; don't just describe them. Batch writes into as few calls as possible (one write_range for a whole block, fill_formula for a calculated column).
- Make output look professional: bold headers, sensible number formats, autofit columns.
- Don't overwrite non-empty cells the user didn't mention; put new results in empty space or a new sheet and say where.
- When explaining a formula or answering a question, be concise and concrete, citing cell addresses.
- Finish with a short summary of what you changed and where.`;

export interface AgentCallbacks {
  onText(delta: string): void;
  onToolStart(name: string, input: unknown): void;
  onToolEnd(name: string, result: unknown, isError: boolean): void;
  onNotice(text: string): void;
}

const MAX_TURNS = 40;

/**
 * Runs one user request to completion: streams Claude's reply, executes the
 * tool calls it makes against the workbook, and loops until Claude is done.
 * `history` is mutated in place so follow-up questions keep context.
 */
export async function runAgent(
  history: MessageParam[],
  userText: string,
  settings: Settings,
  toolCtx: ToolContext,
  cb: AgentCallbacks,
  signal: AbortSignal,
): Promise<void> {
  const client = createClient(settings);
  const tools = activeTools(settings);
  const byName = new Map<string, ToolDef>(tools.map((t) => [t.name, t]));
  const apiTools: Anthropic.Beta.BetaToolUnion[] = tools.map((t) => ({
    name: t.name,
    description: t.description,
    input_schema: t.input_schema as Anthropic.Beta.BetaTool.InputSchema,
    eager_input_streaming: true,
  }));
  if (settings.webSearch) apiTools.push({ type: "web_search_20260209", name: "web_search", max_uses: 5 });
  // Cache the (large, stable) tool list + system prompt across turns.
  const last = apiTools[apiTools.length - 1] as { cache_control?: unknown };
  last.cache_control = { type: "ephemeral" };

  history.push({ role: "user", content: userText });
  let jsonRetries = 0;

  for (let turn = 0; turn < MAX_TURNS; turn++) {
    const stream = client.beta.messages.stream(
      {
        model: settings.model,
        max_tokens: 64000,
        system: [{ type: "text", text: SYSTEM_PROMPT }],
        tools: apiTools,
        messages: history,
        ...requestExtras(settings),
      },
      { signal },
    );
    stream.on("text", (delta) => cb.onText(delta));

    let message: Anthropic.Beta.BetaMessage;
    try {
      message = await stream.finalMessage();
      jsonRetries = 0;
    } catch (err) {
      // Eager tool-input streaming can produce unparseable JSON; re-issue the turn.
      // Anything from the API itself (auth, rate limits, aborts) is rethrown.
      if (err instanceof Anthropic.APIError || signal.aborted || jsonRetries++ >= 2) throw err;
      cb.onNotice("A tool call came back malformed; retrying…");
      continue;
    }

    if (message.stop_reason === "refusal") {
      cb.onNotice("Claude declined this request.");
      return;
    }
    if (message.stop_reason === "pause_turn") {
      history.push({ role: "assistant", content: message.content });
      continue;
    }

    const toolUses = message.content.filter((b): b is Anthropic.Beta.BetaToolUseBlock => b.type === "tool_use");
    if (toolUses.length === 0) {
      history.push({ role: "assistant", content: message.content });
      if (message.stop_reason === "max_tokens") cb.onNotice("Reply was cut off (output limit reached).");
      return;
    }
    if (message.stop_reason === "max_tokens") {
      // A truncated tool input can still parse; never run it.
      history.push({ role: "assistant", content: message.content });
      history.push({
        role: "user",
        content: toolUses.map((t) => ({
          type: "tool_result" as const,
          tool_use_id: t.id,
          is_error: true,
          content: "Tool input was cut off by the output limit. Split the work into smaller calls.",
        })),
      });
      continue;
    }

    history.push({ role: "assistant", content: message.content });
    const results: Anthropic.Beta.BetaToolResultBlockParam[] = [];
    // Run sequentially: Excel operations depend on each other's effects.
    for (const use of toolUses) {
      const tool = byName.get(use.name);
      cb.onToolStart(use.name, use.input);
      let content: string;
      let isError = false;
      if (!tool) {
        content = `Unknown tool ${use.name}`;
        isError = true;
      } else {
        const problems = validate(use.input, tool.input_schema);
        if (problems.length) {
          content = `Invalid input, nothing was changed: ${problems.slice(0, 5).join("; ")}. Raw input: ${JSON.stringify(use.input).slice(0, 2000)}`;
          isError = true;
        } else {
          try {
            content = JSON.stringify(await tool.run(use.input, toolCtx));
          } catch (err) {
            content = describeExcelError(err);
            isError = true;
          }
        }
      }
      cb.onToolEnd(use.name, content, isError);
      results.push({ type: "tool_result", tool_use_id: use.id, content: truncate(content), is_error: isError });
      if (signal.aborted) break;
    }
    // Every tool_use needs a result, even if the user stopped midway.
    for (const use of toolUses.slice(results.length)) {
      results.push({ type: "tool_result", tool_use_id: use.id, content: "Stopped by user.", is_error: true });
    }
    history.push({ role: "user", content: results });
    if (signal.aborted) return;
  }
  cb.onNotice("Stopped after many steps. Ask me to continue if there's more to do.");
}

function describeExcelError(err: unknown): string {
  if (typeof OfficeExtension !== "undefined" && err instanceof OfficeExtension.Error) {
    const loc = err.debugInfo?.errorLocation ? ` (at ${err.debugInfo.errorLocation})` : "";
    return `Excel error ${err.code}: ${err.message}${loc}`;
  }
  return err instanceof Error ? err.message : String(err);
}

function truncate(s: string, max = 60000): string {
  return s.length > max ? `${s.slice(0, max)}… [truncated ${s.length - max} chars]` : s;
}
