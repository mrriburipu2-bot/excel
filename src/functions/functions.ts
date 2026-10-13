/* global CustomFunctions */
// =CLAUDE.* worksheet functions. Registered automatically from the JSDoc
// @customfunction tags by custom-functions-metadata-plugin (namespace CLAUDE
// comes from the manifest). They share settings with the task pane.

import type Anthropic from "@anthropic-ai/sdk";
import { createClient, describeError, requestExtras } from "../shared/claude";
import { loadSettings } from "../shared/settings";

const CELL_SYSTEM =
  "You are a function inside an Excel cell. Reply with only the requested value: no preamble, no explanation, no Markdown, no surrounding quotes. Keep it short enough to fit in a cell unless asked otherwise.";

// Excel recalculates often; reuse answers for identical inputs instead of paying for them again.
const cache = new Map<string, Promise<unknown>>();

function memo<T>(key: string, compute: () => Promise<T>): Promise<T> {
  const hit = cache.get(key);
  if (hit) return hit as Promise<T>;
  const p = compute().catch((err) => {
    cache.delete(key); // don't cache failures
    throw err;
  });
  cache.set(key, p);
  if (cache.size > 2000) cache.delete(cache.keys().next().value as string);
  return p;
}

function rangeToText(values: unknown[][] | undefined): string {
  if (!values) return "";
  return values
    .filter((row) => row.some((v) => v !== "" && v !== null))
    .map((row) => row.map((v) => String(v ?? "")).join("\t"))
    .join("\n");
}

function fail(err: unknown): never {
  throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, describeError(err).slice(0, 250));
}

async function complete(prompt: string, maxTokens = 2048): Promise<string> {
  const settings = loadSettings();
  return memo(`${settings.model}|${prompt}`, async () => {
    try {
      const client = createClient(settings);
      const res = await client.beta.messages.create({
        model: settings.model,
        max_tokens: maxTokens,
        system: CELL_SYSTEM,
        messages: [{ role: "user", content: prompt }],
        ...requestExtras({ ...settings, effort: "low" }),
      });
      if (res.stop_reason === "refusal") return "#REFUSED";
      const text = res.content
        .filter((b): b is Anthropic.Beta.BetaTextBlock => b.type === "text")
        .map((b) => b.text)
        .join("")
        .trim();
      return text;
    } catch (err) {
      fail(err);
    }
  });
}

/**
 * Ask Claude anything. Optionally pass cells as context.
 * @customfunction
 * @param prompt Question or instruction.
 * @param [context] Cells to include as context.
 * @returns Claude's answer.
 */
export async function ask(prompt: string, context?: any[][]): Promise<string> {
  const ctx = rangeToText(context);
  return complete(ctx ? `${prompt}\n\nData (tab-separated):\n${ctx}` : prompt);
}

/**
 * Write an Excel formula from a plain-language description.
 * @customfunction
 * @param description What the formula should do, e.g. "sum column B where column A is North".
 * @param [sample] Optional sample of the data (with headers) the formula will work on.
 * @returns The formula as text.
 */
export async function formula(description: string, sample?: any[][]): Promise<string> {
  const ctx = rangeToText(sample);
  return complete(
    `Write one Excel formula (English function names, comma separators) that does this: ${description}` +
      (ctx ? `\n\nSample data (tab-separated, first row is headers):\n${ctx}` : "") +
      "\nReturn only the formula starting with =.",
  );
}

/**
 * Classify text into one of the given categories.
 * @customfunction
 * @param text Text to classify.
 * @param categories Cells listing the categories (or one cell with a comma-separated list).
 * @returns The best-matching category.
 */
export async function classify(text: string, categories: string[][]): Promise<string> {
  const list = categories.flat().filter(Boolean).join(", ");
  return complete(`Classify the text into exactly one of these categories: ${list}.\nText: ${text}\nReturn only the category.`, 256);
}

/**
 * Extract a piece of information from text (e.g. email, company, amount, date).
 * @customfunction
 * @param text Source text.
 * @param what What to extract.
 * @returns The extracted value, or empty if not found.
 */
export async function extract(text: string, what: string): Promise<string> {
  return complete(`Extract the ${what} from the text below. Return only the value, or an empty string if it's not present.\nText: ${text}`, 512);
}

/**
 * Translate text into another language.
 * @customfunction
 * @param text Text to translate.
 * @param language Target language, e.g. "Spanish".
 * @returns Translated text.
 */
export async function translate(text: string, language: string): Promise<string> {
  return complete(`Translate into ${language}:\n${text}`);
}

/**
 * Summarize text or a range of cells.
 * @customfunction
 * @param data Cells to summarize.
 * @param [instructions] Optional focus, e.g. "in 10 words" or "key risks".
 * @returns Summary.
 */
export async function summarize(data: any[][], instructions?: string): Promise<string> {
  return complete(`Summarize this data${instructions ? ` (${instructions})` : ""}:\n${rangeToText(data)}`);
}

/**
 * Sentiment of a piece of text: Positive, Negative, Neutral or Mixed.
 * @customfunction
 * @param text Text to analyze.
 * @returns Sentiment label.
 */
export async function sentiment(text: string): Promise<string> {
  return complete(`What is the sentiment of this text? Answer with exactly one of: Positive, Negative, Neutral, Mixed.\nText: ${text}`, 64);
}

/**
 * Generate a table of data that spills into neighbouring cells.
 * @customfunction
 * @param prompt What table to generate, e.g. "10 largest countries by area with capital and population".
 * @returns Table with a header row.
 */
export async function table(prompt: string): Promise<string[][]> {
  const settings = loadSettings();
  return memo(`table|${settings.model}|${prompt}`, async () => {
    try {
      const client = createClient(settings);
      const extras = requestExtras({ ...settings, effort: "low" });
      const res = await client.beta.messages.create({
        model: settings.model,
        max_tokens: 8192,
        system: "You generate tabular data for Excel. The first row is the header row. Every row has the same number of columns.",
        messages: [{ role: "user", content: prompt }],
        betas: extras.betas,
        ...(extras.fallbacks ? { fallbacks: extras.fallbacks } : {}),
        output_config: {
          ...extras.output_config,
          format: {
            type: "json_schema",
            schema: {
              type: "object",
              properties: { rows: { type: "array", items: { type: "array", items: { type: "string" } } } },
              required: ["rows"],
              additionalProperties: false,
            },
          },
        },
      });
      if (res.stop_reason === "refusal") return [["#REFUSED"]];
      const text = res.content.find((b): b is Anthropic.Beta.BetaTextBlock => b.type === "text")?.text ?? "{}";
      const rows: string[][] = JSON.parse(text).rows ?? [];
      const width = Math.max(1, ...rows.map((r) => r.length));
      return rows.length ? rows.map((r) => [...r, ...Array(width - r.length).fill("")]) : [[""]];
    } catch (err) {
      fail(err);
    }
  });
}
