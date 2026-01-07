# XLify — ChatGPT‑powered Excel Add‑in

XLify is a VSTO (Visual Studio Tools for Office) add‑in that lets you control Excel through natural‑language chat. It uses Semantic Kernel for tool/function calling, a WebView2 UI for chat, and an out‑of‑process Roslyn C# executor to run generated automation code safely.

## Highlights

- Chat with Excel: ask for summaries, populate data, format sheets, and more.
- Tool‑calling: the model calls a single tool `roslyn.execute_csharp` to run C# script against Excel.
- Safety + reliability: code runs in a separate COM local server (out of process) that hosts Roslyn, isolating Excel from compiler crashes.
- Observability: rich, structured logs (tools, prompts, results, timings) streamed to Seq.
- Responsive UI: “Thinking…” vs “Working…” live status, auto‑scrolling chat, auto‑expanding composer.

## Repository layout

- `XLify/` — VSTO Excel add‑in (C#, .NET Framework 4.8)
  - Hosts WebView2 UI and wires Semantic Kernel
  - Adds SK plugins (Roslyn tool, documentation helpers)
  - Status bridge to WebView2
  - Logging to Seq
- `XLify.CSharpExecutor/` — out‑of‑process COM server (C#, .NET Framework 4.8)
  - Roslyn C# scripting engine (Microsoft.CodeAnalysis.CSharp.Scripting)
  - Per‑session script state reuse; Excel `Application` injected via globals
  - COM visible (`ProgId: XLify.CSharpExecutor`), registered automatically in Debug via `regasm`
- `WebApp/` — WebView2 front‑end (React + Vite)
  - Minimal “assistant UI”: messages list, status bubble, auto‑growing composer, Enter‑to‑send
  - Listens to host messages via `window.chrome.webview`

## Architecture

```
User ↔ WebView2 (WebApp) ↔ VSTO Add‑in (XLify) ↔ Semantic Kernel ↔ OpenAI
                                       │
                                       └─ Tool: roslyn.execute_csharp(code)
                                             ↕  COM (local server)
                                        CSharpExecutor (Roslyn) ↔ Excel COM
```

1) The user types in the WebView UI; the add‑in receives messages and builds a chat history + system instructions.

2) Semantic Kernel (SK) sends the prompt to OpenAI. When the model requests a tool call, SK invokes the `roslyn.execute_csharp` tool with the generated C# code.

3) The add‑in calls the COM local server `XLify.CSharpExecutor`, which hosts Roslyn to compile/execute the script against Excel’s COM object model.

4) The tool returns stdout/stderr back to the add‑in, which streams the result to the WebView chat.

## Key components

### Semantic Kernel setup

- SK chat completion using OpenAI (`Microsoft.SemanticKernel.Connectors.OpenAI`).
- Plugins:
  - `roslyn.execute_csharp` — executes C# scripts in the current Excel session (see Roslyn tool).
  - `doc.*` — helper functions for workbook/selection summaries.
- Filters & logging:
  - Function invocation filter (timing + args + result): logs `SK-TOOL CALL/DONE/ERROR`.
  - Prompt render filter: logs rendered prompts (`SK-PROMPT RENDER/RENDERED`).
  - Message‑level tool calls/results (streamed): logs `SK-MSG TOOL-CALL/RESULT`.
- Performance guardrails baked into the system message (batch operations with `object[,]`, suspend UI/calculation, avoid AutoFit in loops, etc.).

### Roslyn C# executor (COM local server)

- Exposes `ICSharpExecutor` (CreateSession/Reset/Destroy/ExecuteInSession/ExecuteOneOff).
- Injects globals (`ExcelApp`, and `ApplicationInstance`) and bootstraps code so scripts can use:
  - `var Application = (ApplicationInstance ?? ExcelApp);`
  - `using Excel = Microsoft.Office.Interop.Excel;` (alias pre‑imported for enums like `Excel.XlCalculation.*`).
- Reuses script state per session to amortize compilation/JIT across calls.
- Captures and returns stdout, stderr, and return values.
- Writes structured logs to Seq (`Subsystem=Worker`, `App=XLify.CSharpExecutor`).

### WebView2 front‑end

- React UI with a simple chat layout and status bubble.
- Auto‑scroll to bottom on new messages/status.
- Auto‑expanding textarea (Enter=send, Shift+Enter=newline).
- Listens for messages from host:
  - `type: 'assistant' | 'user' | 'error' | 'debug' | 'status'`.

## Observability (Seq)

- Serilog sinks to Seq are configured in both the add‑in and the worker.
- Tagging:
  - Add‑in: `App=XLify.AddIn`, `Subsystem=SK` for SK logs, `Workspace=XLify-AddIn`.
  - Worker: `App=XLify.CSharpExecutor`, `Subsystem=Worker`, `Workspace=XLify-Worker`.
- Workspaces: create two in Seq with default filters `App = 'XLify.AddIn'` and `App = 'XLify.CSharpExecutor'` for clean separation (workspaces are filtered views, not separate storage).
- Code privacy: by default, logs include only code length + SHA‑1, not full content. Set `XLIFY_LOG_CODE=1` to include truncated code for diagnostics.
- Environment variables:
  - `SEQ_URL` (e.g., `http://localhost:5341`)
  - `SEQ_API_KEY` (optional; if your Seq requires a key)
  - `XLIFY_LOG_CODE` (`1`/`true` to log code content)

## Performance design

- Warm‑up: a one‑time Roslyn+Excel warm‑up runs after session creation to reduce first‑call latency.
- Guidance for generated code (system prompt):
  - Batch updates with `object[,]` and single `Range.Value2` assignment.
  - Wrap heavy work with `ScreenUpdating=false`, `DisplayAlerts=false`, and `Calculation = Excel.XlCalculation.xlCalculationManual` (then restore).
  - Avoid whole‑sheet operations and per‑cell loops; set `ColumnWidth` once instead of `AutoFit` in loops.
  - Use `dynamic` or explicit casts for COM objects; never store COM proxies in `object`.

## Building & running

Prerequisites:

- Visual Studio with Office/Excel developer tools
- .NET Framework 4.8 targeting packs
- Excel (desktop)
- WebView2 Runtime (Evergreen)

Steps:

1) Build the solution `XLify.sln` in Debug.
   - The COM server (`XLify.CSharpExecutor`) registers per‑user via `regasm` post‑build in Debug.
2) Launch Excel; the XLify ribbon appears, and you can open the task pane (WebView UI).
3) Enter your OpenAI API key in the UI (stored locally via `ApiKeyVault`).
4) Chat with Excel. The add‑in auto‑creates a Roslyn session and warms up.

Notes:

- If Excel cannot create the COM server, ensure the COM registration target ran (build in Debug) or register manually.
- Ensure Seq is running at `SEQ_URL` if you want logs; otherwise, logging calls are harmless.

## Troubleshooting

- "COM server not registered": build in Debug or register `XLify.CSharpExecutor.exe` with `regasm` (x86/x64 matching your Excel).
- Compile errors like `CS0266` or `CS1061` in scripts:
  - Use `dynamic` for COM proxies or cast to `Excel.Range` before member access.
  - Avoid storing COM objects in variables typed as `object`.
  - Prefer `Excel.XlCalculation.*` enums over raw integers.
- Slow runs: confirm the code uses `object[,]` bulk writes, and the prompt includes performance instructions.

## Status integration

- The add‑in sends live status to the WebView:
  - `"Thinking…"` while waiting on the model/API
  - `"Working…"` while executing code via the tool
  - status cleared when the assistant message is shown

## Security

- Generated C# runs with the privileges of the Excel process. Treat API keys and workbooks as sensitive, and consider environment isolation for untrusted code.

## License

Proprietary / internal use (no explicit open‑source license is included in this repository).

