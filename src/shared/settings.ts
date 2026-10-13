// User settings, persisted in localStorage. The task pane and the custom
// functions run in the same shared runtime, so both see the same values.

export type Effort = "low" | "medium" | "high" | "xhigh" | "max";

export interface Settings {
  apiKey: string;
  model: string;
  effort: Effort;
  /** Optional proxy that forwards to the Anthropic API (keeps the key off the client). */
  baseURL: string;
  webSearch: boolean;
  /** Lets the assistant run arbitrary Office.js code (each run still needs approval). */
  allowCode: boolean;
}

export const MODELS: { id: string; label: string }[] = [
  { id: "claude-opus-5-5", label: "Claude Opus 5.5 (recommended)" },
  { id: "claude-fable-5-1", label: "Claude Fable 5.1 (most capable)" },
  { id: "claude-sonnet-5-5", label: "Claude Sonnet 5.5" },
  { id: "claude-haiku-5-5", label: "Claude Haiku 5.5 (fastest)" },
];

export const DEFAULT_SETTINGS: Settings = {
  apiKey: "",
  model: "claude-opus-5-5",
  effort: "medium",
  baseURL: "",
  webSearch: false,
  allowCode: false,
};

const KEY = "claude-excel-settings";

function storage(): Storage | null {
  try {
    return window.localStorage;
  } catch {
    return null;
  }
}

export function loadSettings(): Settings {
  try {
    const raw = storage()?.getItem(KEY);
    return raw ? { ...DEFAULT_SETTINGS, ...JSON.parse(raw) } : { ...DEFAULT_SETTINGS };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: Settings): void {
  try {
    storage()?.setItem(KEY, JSON.stringify(settings));
  } catch {
    // Storage blocked (private mode, policy); settings last for this session only.
  }
}
