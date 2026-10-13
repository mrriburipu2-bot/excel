import Anthropic from "@anthropic-ai/sdk";
import type { Settings } from "./settings";

export class MissingApiKeyError extends Error {
  constructor() {
    super("No Anthropic API key set. Open the AI Assistant task pane and add one in Settings.");
  }
}

export function createClient(settings: Settings): Anthropic {
  if (!settings.apiKey && !settings.baseURL) throw new MissingApiKeyError();
  return new Anthropic({
    apiKey: settings.apiKey || "proxy",
    baseURL: settings.baseURL || undefined,
    defaultHeaders: settings.workspaceId ? { "anthropic-workspace-id": settings.workspaceId } : undefined,
    // The add-in runs in the Office webview, so the SDK is used from a browser.
    // The key is the user's own and stays in their local storage; use baseURL
    // with a server-side proxy to keep keys off client machines entirely.
    dangerouslyAllowBrowser: true,
  });
}

// Models that accept server-side refusal fallbacks with `fallbacks: "default"`.
const FALLBACK_MODELS = new Set(["claude-fable-5-1", "claude-opus-5-5", "claude-opus-5", "claude-sonnet-5-5"]);

/** Extra request fields shared by every call: effort plus refusal fallbacks where supported. */
export function requestExtras(settings: Settings): {
  betas: Anthropic.Beta.AnthropicBeta[];
  fallbacks?: "default";
  output_config: { effort: Settings["effort"] };
} {
  const useFallbacks = FALLBACK_MODELS.has(settings.model);
  return {
    betas: useFallbacks ? ["server-side-fallback-2026-07-01"] : [],
    ...(useFallbacks ? { fallbacks: "default" as const } : {}),
    output_config: { effort: settings.effort },
  };
}

export function describeError(err: unknown): string {
  if (err instanceof MissingApiKeyError) return err.message;
  if (err instanceof Anthropic.AuthenticationError) return "Invalid API key. Check it in Settings.";
  if (err instanceof Anthropic.PermissionDeniedError) return "This API key is not allowed to use the selected model.";
  if (err instanceof Anthropic.NotFoundError) return "Model not found. Pick another model in Settings.";
  if (err instanceof Anthropic.RateLimitError) return "Rate limited by the Anthropic API. Wait a moment and try again.";
  if (err instanceof Anthropic.APIConnectionError) return "Could not reach the Anthropic API. Check your network or proxy URL.";
  if (err instanceof Anthropic.BadRequestError && /anthropic-workspace-id|scoped to a workspace/i.test(err.message)) {
    return "This API key isn't tied to a workspace. Open Settings and enter your Workspace ID (Claude Console → Settings → Workspaces), or create a key inside a workspace.";
  }
  if (err instanceof Anthropic.APIError) return `Anthropic API error ${err.status ?? ""}: ${err.message}`;
  if (err instanceof Error) return err.message;
  return String(err);
}
