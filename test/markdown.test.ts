import { test } from "node:test";
import assert from "node:assert/strict";
import { renderMarkdown } from "../src/shared/markdown.ts";

test("escapes HTML from model output", () => {
  const html = renderMarkdown('<img src=x onerror="alert(1)">');
  assert.ok(!html.includes("<img"));
  assert.ok(html.includes("&lt;img"));
});

test("renders lists, code, bold and tables", () => {
  const html = renderMarkdown("**Done**\n\n- one\n- `=SUM(A1:A3)`\n\n| A | B |\n|---|---|\n| 1 | 2 |\n\n```\n=IF(A1>0,1,0)\n```");
  assert.ok(html.includes("<strong>Done</strong>"));
  assert.ok(html.includes("<ul><li>one</li><li><code>=SUM(A1:A3)</code></li></ul>"));
  assert.ok(html.includes("<th>A</th>") && html.includes("<td>2</td>"));
  assert.ok(html.includes("<pre><code>=IF(A1&gt;0,1,0)</code></pre>"));
});
