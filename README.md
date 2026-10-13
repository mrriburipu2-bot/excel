# Claude AI for Excel

An Excel add-in that brings Claude into your workbook in two ways:

1. **AI Assistant task pane.** Chat in plain language and Claude works on the open workbook directly. It reads your sheets, writes formulas and data, and formats, sorts, filters, charts, pivots, validates and builds dashboards for you.
2. **`=CLAUDE.*` worksheet functions.** Use AI inside cells, e.g. `=CLAUDE.ASK("Summarize", A2:D20)`.

Works in Excel for Windows, Mac and Excel on the web (Office.js, shared runtime).

## What the assistant can do

The assistant has 25 workbook tools that cover the main areas of Excel:

| Area | Tools |
|---|---|
| Understand the workbook | `get_workbook_info`, `read_range`, `get_selection`, `get_range_statistics` |
| Data entry and formulas | `write_range` (values plus formulas, including dynamic arrays, XLOOKUP, LET, LAMBDA…), `fill_formula`, `copy_range`, `find_replace`, `clear_range` |
| Formatting | `format_range` (font, fill, number formats, alignment, borders, wrap, merge, widths/heights, autofit) |
| Structure | `manage_worksheet` (add/delete/rename/copy/move/hide/tab color/freeze panes/protect), `modify_rows_columns` (insert/delete/hide/group/autofit) |
| Tables and data tools | `create_table`, `modify_table` (style, totals, add rows…), `sort_range`, `filter_data`, `remove_duplicates`, `set_data_validation`, `manage_named_range` |
| Analysis and visuals | `create_chart` (26 chart types), `create_pivot_table`, `add_conditional_format` (color scales, data bars, icon sets, rules, top/bottom, duplicates, formula rules) |
| Misc | `annotate_cell` (comments and hyperlinks), `select_range`, `recalculate` |
| Advanced (opt-in) | `run_office_js` runs Office.js code for anything else, such as slicers, shapes or page setup. It is off by default, and you approve every run. |

Optional **web search** lets Claude look things up, such as exchange rates or definitions.

Quick-action buttons cover common jobs: explain the selection, summarize a sheet, clean data, format as a table, make a chart, build a PivotTable, find formula errors, highlight outliers, build a dashboard, and write a formula.

## Worksheet functions

| Function | Example |
|---|---|
| `CLAUDE.ASK(prompt, [context])` | `=CLAUDE.ASK("Which region grew fastest?", A1:E50)` |
| `CLAUDE.FORMULA(description, [sample])` | `=CLAUDE.FORMULA("sum Sales where Region is North", A1:C5)` |
| `CLAUDE.CLASSIFY(text, categories)` | `=CLAUDE.CLASSIFY(B2, $H$2:$H$6)` |
| `CLAUDE.EXTRACT(text, what)` | `=CLAUDE.EXTRACT(C2, "email address")` |
| `CLAUDE.TRANSLATE(text, language)` | `=CLAUDE.TRANSLATE(A2, "French")` |
| `CLAUDE.SUMMARIZE(data, [instructions])` | `=CLAUDE.SUMMARIZE(A2:A200, "top 3 complaints")` |
| `CLAUDE.SENTIMENT(text)` | `=CLAUDE.SENTIMENT(D2)` returns Positive, Negative, Neutral or Mixed |
| `CLAUDE.TABLE(prompt)` | `=CLAUDE.TABLE("10 largest countries by area with capital")` spills a table |

Results are cached per input, so recalculation doesn't repeat API calls. Cell functions run at low effort to keep them fast and cheap.

## Setup

This add-in isn't in the Office Store, so it won't appear under **Home → Add-ins → My Add-ins**. You have to *sideload* it, and the add-in's files must be running on your PC while you use it.

### Excel for Windows (desktop)

You only need **Node.js** (LTS from [nodejs.org](https://nodejs.org)). No `npm install` is needed: the repository ships a ready-built add-in (`windows/addin`, about 0.5 MB).

1. Download this repository: **Code → Download ZIP** on GitHub, then unzip it.
2. Double-click **`install-windows.cmd`**.
   - If Excel is open, it asks you to save and close it first.
   - Windows asks whether to install a certificate for `localhost`. Click **Yes**. The certificate lets Excel load the add-in securely from your own PC and is installed for your user only.
3. Excel opens with **Home → Claude AI → AI Assistant**. Keep the black window open while you use the add-in; closing it stops the add-in.

Each time you want to use the add-in, double-click `install-windows.cmd` again; later runs skip the setup steps. To remove the add-in and its certificate, run `uninstall-windows.cmd`.

> The Office Add-ins dialog in desktop Excel only lists Store add-ins and has no "Upload My Add-in" option. That's why the script registers the add-in with Excel directly.

Developers: after changing the source, run `npm run package:windows` to rebuild `windows/addin`.

### Excel on the web

```bash
npm install
npm run dev-server     # keep running; serves https://localhost:3000
```

Then in Excel on the web go to **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in** and choose `manifest.xml`.

### Excel for Mac / development

```bash
npm install            # full developer tooling (several hundred MB)
npm start              # starts the dev server and sideloads into Excel desktop
```

### First use

1. Click **Home → AI Assistant**.
2. Paste your Anthropic API key (from [console.anthropic.com](https://console.anthropic.com)) in Settings, then click Save.
3. Ask something, for example "Add a profit column, format it as currency and chart profit by region".

### Settings

- **Model**: Claude Opus 5.5 (default), Fable 5.1, Sonnet 5.5 or Haiku 5.5.
- **Effort**: from low (fastest) to max (most thorough). Medium is the default.
- **Workspace ID**: only needed if you see "This API key is not scoped to a workspace". Copy it from Claude Console → Settings → Workspaces, or create the API key inside a workspace instead.
- **API proxy URL**: point the add-in at your own server that forwards to the Anthropic API, so API keys never reach users' machines (recommended for organization-wide deployment).
- **Web search** and **advanced Office.js code** are off by default.

## Deploying

1. Set `urlProd` in `webpack.config.js` to your HTTPS host.
2. Run `npm run build` and host the `dist/` folder.
3. Distribute `dist/manifest.xml`, either through the Microsoft 365 admin center (Integrated Apps) or a SharePoint app catalog.

## Development

```bash
npm run typecheck   # TypeScript
npm test            # unit tests (address parsing, stats, schema validation, Markdown rendering)
npm run build       # production bundle in dist/
```

Project layout:

```
manifest.xml              Add-in manifest (task pane, ribbon button, CLAUDE custom functions, shared runtime)
src/taskpane/agent.ts     Agent loop: streams Claude, validates and runs tool calls, loops until done
src/taskpane/tools.ts     All workbook tools (schemas + Office.js implementations)
src/taskpane/taskpane.*   Chat UI, settings, approvals
src/functions/functions.ts  =CLAUDE.* custom functions
src/shared/               Settings, API client, schema validator, Excel helpers, Markdown renderer
```

## Security notes

- The API key is stored in the add-in's local storage on the user's machine and sent only to the Anthropic API (or to your proxy).
- Model output is rendered as escaped text. The assistant can't inject HTML into the task pane.
- Tool inputs are validated against their schemas before they touch the workbook. Deleting a sheet and running custom code always require your approval.
- Office.js changes made by add-ins can't be undone with Ctrl+Z, so keep a copy of important workbooks.
