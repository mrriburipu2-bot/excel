import type Anthropic from "@anthropic-ai/sdk";
import { describeError } from "../shared/claude";
import { renderMarkdown } from "../shared/markdown";
import { type Effort, loadSettings, MODELS, saveSettings, type Settings } from "../shared/settings";
import { runAgent } from "./agent";

const QUICK_ACTIONS: { label: string; prompt: string }[] = [
  { label: "Explain selection", prompt: "Explain what the selected cells contain and how any formulas in them work." },
  { label: "Summarize this sheet", prompt: "Summarize the data on the active sheet: what it contains, key totals, trends and anything unusual." },
  { label: "Clean data", prompt: "Clean the data on the active sheet: trim spaces, fix inconsistent casing and date/number formats, remove duplicate rows, and tell me what you changed." },
  { label: "Format as table", prompt: "Turn the data on the active sheet into a nicely formatted Excel table with proper number formats and autofit columns." },
  { label: "Make a chart", prompt: "Create the most useful chart for the data on the active sheet and place it next to the data." },
  { label: "Pivot table", prompt: "Build a PivotTable on a new sheet that summarizes the main numeric column by the most useful category." },
  { label: "Find formula errors", prompt: "Check the workbook for formula errors (#N/A, #REF!, #DIV/0!, #VALUE!, inconsistent formulas) and fix them or explain how." },
  { label: "Highlight outliers", prompt: "Find outliers in the numeric columns of the active sheet and highlight them with conditional formatting." },
  { label: "Build a dashboard", prompt: "Build a dashboard sheet with KPIs, a summary table and charts based on the data in this workbook." },
  { label: "Write a formula", prompt: "Help me write a formula. Ask me what I need if it isn't clear from the selection." },
];

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

let settings: Settings = loadSettings();
let history: Anthropic.Beta.BetaMessageParam[] = [];
let controller: AbortController | null = null;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    appendMessage("error", "This add-in runs in Excel.");
    return;
  }
  initSettings();
  initQuickActions();
  $("send").addEventListener("click", onSendClick);
  $<HTMLTextAreaElement>("prompt").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      onSendClick();
    }
  });
  $("new-chat").addEventListener("click", () => {
    if (controller) return;
    history = [];
    $("log").innerHTML = "";
    $("welcome").hidden = false;
  });
  if (!settings.apiKey && !settings.baseURL) showSettings();
});

// ---------- settings ----------

function initSettings() {
  const model = $<HTMLSelectElement>("model");
  for (const m of MODELS) model.add(new Option(m.label, m.id));
  $("open-settings").addEventListener("click", showSettings);
  $("close-settings").addEventListener("click", () => ($("settings").hidden = true));
  $("save-settings").addEventListener("click", () => {
    settings = {
      apiKey: $<HTMLInputElement>("api-key").value.trim(),
      model: model.value,
      effort: $<HTMLSelectElement>("effort").value as Effort,
      baseURL: $<HTMLInputElement>("base-url").value.trim(),
      webSearch: $<HTMLInputElement>("web-search").checked,
      allowCode: $<HTMLInputElement>("allow-code").checked,
    };
    saveSettings(settings);
    $("settings").hidden = true;
  });
}

function showSettings() {
  $<HTMLInputElement>("api-key").value = settings.apiKey;
  $<HTMLSelectElement>("model").value = settings.model;
  $<HTMLSelectElement>("effort").value = settings.effort;
  $<HTMLInputElement>("base-url").value = settings.baseURL;
  $<HTMLInputElement>("web-search").checked = settings.webSearch;
  $<HTMLInputElement>("allow-code").checked = settings.allowCode;
  $("settings").hidden = false;
}

function initQuickActions() {
  const box = $("quick-actions");
  for (const a of QUICK_ACTIONS) {
    const b = document.createElement("button");
    b.textContent = a.label;
    b.addEventListener("click", () => void send(a.prompt));
    box.appendChild(b);
  }
}

// ---------- chat ----------

function onSendClick() {
  if (controller) {
    controller.abort();
    return;
  }
  const box = $<HTMLTextAreaElement>("prompt");
  const text = box.value.trim();
  if (!text) return;
  box.value = "";
  void send(text);
}

async function send(text: string) {
  if (controller) return;
  if (!settings.apiKey && !settings.baseURL) {
    showSettings();
    return;
  }
  $("welcome").hidden = true;
  appendMessage("user", text);
  controller = new AbortController();
  setBusy(true);

  let bubble: HTMLElement | null = null;
  let buffer = "";
  const flush = () => {
    if (bubble) bubble.innerHTML = renderMarkdown(buffer);
  };

  try {
    const selection = await currentSelection();
    const userText = selection ? `${text}\n\n(Current selection: ${selection})` : text;
    await runAgent(
      history,
      userText,
      settings,
      { settings, confirm: askApproval },
      {
        onText(delta) {
          if (!bubble) {
            bubble = appendMessage("assistant", "");
            bubble.classList.add("typing");
            buffer = "";
          }
          buffer += delta;
          flush();
          scrollDown();
        },
        onToolStart(name, input) {
          if (bubble) bubble.classList.remove("typing");
          bubble = null;
          appendTool(name, input);
        },
        onToolEnd(name, result, isError) {
          finishTool(name, result, isError);
        },
        onNotice(t) {
          appendMessage("notice", t);
        },
      },
      controller.signal,
    );
  } catch (err) {
    if (!controller.signal.aborted) appendMessage("error", describeError(err));
    else appendMessage("notice", "Stopped.");
    // Drop the unfinished exchange so the next request starts from a valid history.
    trimHistory();
  } finally {
    (bubble as HTMLElement | null)?.classList.remove("typing");
    controller = null;
    setBusy(false);
  }
}

/** Removes trailing turns that would leave the conversation in an invalid state. */
function trimHistory() {
  while (history.length) {
    const last = history[history.length - 1];
    const endsWithToolUse =
      last.role === "assistant" && Array.isArray(last.content) && last.content.some((b) => b.type === "tool_use");
    if (last.role === "user" || endsWithToolUse) history.pop();
    else break;
  }
}

async function currentSelection(): Promise<string | null> {
  try {
    return await Excel.run(async (context) => {
      const r = context.workbook.getSelectedRange();
      r.load("address");
      await context.sync();
      return r.address;
    });
  } catch {
    return null;
  }
}

function setBusy(busy: boolean) {
  const b = $<HTMLButtonElement>("send");
  b.textContent = busy ? "Stop" : "Send";
  b.classList.toggle("primary", !busy);
  $<HTMLButtonElement>("new-chat").disabled = busy;
}

function appendMessage(kind: "user" | "assistant" | "notice" | "error", text: string): HTMLElement {
  const div = document.createElement("div");
  div.className = `msg ${kind}`;
  if (kind === "assistant") div.innerHTML = renderMarkdown(text);
  else div.textContent = text;
  $("log").appendChild(div);
  scrollDown();
  return div;
}

const TOOL_LABELS: Record<string, string> = {
  get_workbook_info: "Looking at the workbook",
  read_range: "Reading cells",
  get_selection: "Reading selection",
  write_range: "Writing cells",
  fill_formula: "Filling formula",
  format_range: "Formatting",
  clear_range: "Clearing cells",
  modify_rows_columns: "Changing rows/columns",
  manage_worksheet: "Updating worksheet",
  create_table: "Creating table",
  modify_table: "Updating table",
  sort_range: "Sorting",
  filter_data: "Filtering",
  create_chart: "Creating chart",
  add_conditional_format: "Adding conditional format",
  set_data_validation: "Setting data validation",
  manage_named_range: "Updating named range",
  create_pivot_table: "Creating PivotTable",
  find_replace: "Find / replace",
  remove_duplicates: "Removing duplicates",
  copy_range: "Copying cells",
  annotate_cell: "Adding comment/link",
  get_range_statistics: "Analyzing data",
  select_range: "Selecting",
  recalculate: "Recalculating",
  run_office_js: "Running code",
};

let openTool: HTMLDetailsElement | null = null;

function appendTool(name: string, input: unknown) {
  const d = document.createElement("details");
  d.className = "tool";
  const s = document.createElement("summary");
  const target = (input as { address?: string })?.address ?? "";
  s.textContent = `⏳ ${TOOL_LABELS[name] ?? name}${target ? ` · ${target}` : ""}`;
  const pre = document.createElement("pre");
  pre.textContent = JSON.stringify(input, null, 2);
  d.append(s, pre);
  $("log").appendChild(d);
  openTool = d;
  scrollDown();
}

function finishTool(name: string, result: unknown, isError: boolean) {
  if (!openTool) return;
  const s = openTool.querySelector("summary")!;
  s.textContent = s.textContent!.replace("⏳", isError ? "⚠️" : "✓");
  if (isError) openTool.classList.add("err");
  const pre = document.createElement("pre");
  pre.textContent = String(result).slice(0, 4000);
  openTool.appendChild(pre);
  openTool = null;
}

function askApproval(title: string, detail: string): Promise<boolean> {
  return new Promise((resolve) => {
    const box = $("confirm");
    $("confirm-title").textContent = title;
    $("confirm-detail").textContent = detail;
    box.hidden = false;
    const done = (ok: boolean) => {
      box.hidden = true;
      $("confirm-yes").removeEventListener("click", yes);
      $("confirm-no").removeEventListener("click", no);
      resolve(ok);
    };
    const yes = () => done(true);
    const no = () => done(false);
    $("confirm-yes").addEventListener("click", yes);
    $("confirm-no").addEventListener("click", no);
    // Pressing Stop while a request waits for approval counts as Deny.
    controller?.signal.addEventListener("abort", no, { once: true });
  });
}

function scrollDown() {
  const main = $("chat");
  main.scrollTop = main.scrollHeight;
}
