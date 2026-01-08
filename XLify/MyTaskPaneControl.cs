using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Net;
 
using System.Threading;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion; // still used for ChatHistory and message types
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Microsoft.SemanticKernel.Agents;
using Microsoft.SemanticKernel.Agents.OpenAI;
using OpenAI;
using OpenAI.Responses;
using Serilog;
 

namespace XLify
{
    public class MyTaskPaneControl : UserControl
    {
        private WebView2 _web;
        private readonly string _sessionId = Guid.NewGuid().ToString("n");
        private bool _sessionCreated;
        // Removed: direct HTTP client for Responses; rely on SK agent
        private Kernel _kernel;
        // Removed: SK IChatCompletionService; we use the Responses agent only
        private OpenAIResponseAgent _responseAgent;
        private readonly SemaphoreSlim _chatLock = new SemaphoreSlim(1, 1);
        private bool _sessionWarmedUp;
        private readonly SynchronizationContext _uiContext;
        // Using ChatClient directly per SDK docs

        public MyTaskPaneControl()
        {
            _uiContext = SynchronizationContext.Current;
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12; } catch { }
            this.Dock = DockStyle.Fill;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Padding = new Padding(0); // no outer padding to avoid gap around WebView
            this.Margin = new Padding(0);

            _web = new WebView2
            {
                Dock = DockStyle.Fill,
                AllowExternalDrop = false,
                DefaultBackgroundColor = Color.White,
                Margin = new Padding(0),
            };
            this.Controls.Add(_web);

            this.Load += MyTaskPaneControl_Load;
        }

        private async void MyTaskPaneControl_Load(object sender, EventArgs e)
        {
            await InitializeWebViewAsync();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
            }
            base.Dispose(disposing);
        }

        private async Task InitializeWebViewAsync()
        {
            try
            {
                // Ensure the environment (uses installed Evergreen runtime) with explicit user data folder
                var userData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XLify", "WebView2UserData");
                Directory.CreateDirectory(userData);
                var env = await CoreWebView2Environment.CreateAsync(null, userData);
                await _web.EnsureCoreWebView2Async(env);

                // Basic settings
                _web.CoreWebView2.Settings.AreDevToolsEnabled = true;
                _web.CoreWebView2.Settings.IsScriptEnabled = true;
                _web.CoreWebView2.Settings.IsZoomControlEnabled = true;

                _web.CoreWebView2.WebMessageReceived += async (s, e) =>
                {
                    var json = e.WebMessageAsJson; // full JSON (strings are quoted JSON)
                    await OnUiAsync(async () =>
                    {
                        try
                        {
                            try { if (ShouldLogWebView2Debug()) Debug.WriteLine($"[WebView2] {json}"); } catch { }
                            try { Console.WriteLine($"[WebView2] {json}"); } catch { }

                        // Extract text if present
                        string inputText = null;
                        string messageType = null;
                        string providedKey = null;
                        try
                        {
                            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                            var parsed = serializer.DeserializeObject(json);
                            if (parsed is string sstr)
                            {
                                inputText = sstr;
                            }
                            else if (parsed is System.Collections.Generic.Dictionary<string, object> dict)
                            {
                                if (dict.ContainsKey("type")) messageType = dict["type"] as string;
                                if (dict.ContainsKey("text")) inputText = dict["text"] as string;
                                if (dict.ContainsKey("apiKey")) providedKey = dict["apiKey"] as string;
                            }
                        }
                        catch { }

                        // Handle key management messages
                        if (string.Equals(messageType, "hasApiKey", StringComparison.OrdinalIgnoreCase))
                        {
                            var resp = BuildJsonSafe("hasApiKey", null, ApiKeyVault.Has() ? "true" : "false");
                            OnUi(() => _web.CoreWebView2.PostWebMessageAsString(resp));
                            return;
                        }
                        if (string.Equals(messageType, "saveApiKey", StringComparison.OrdinalIgnoreCase))
                        {
                            try { if (!string.IsNullOrWhiteSpace(providedKey)) ApiKeyVault.Save(providedKey); } catch { }
                            var resp = BuildJsonSafe("saveApiKey", null, "ok");
                            OnUi(() => _web.CoreWebView2.PostWebMessageAsString(resp));
                            return;
                        }
                        if (string.Equals(messageType, "clearApiKey", StringComparison.OrdinalIgnoreCase))
                        {
                            ApiKeyVault.Clear();
                            var resp = BuildJsonSafe("clearApiKey", null, "ok");
                            OnUi(() => _web.CoreWebView2.PostWebMessageAsString(resp));
                            return;
                        }

                        if ((string.IsNullOrWhiteSpace(messageType) || string.Equals(messageType, "user", StringComparison.OrdinalIgnoreCase)) && !string.IsNullOrWhiteSpace(inputText))
                        {
                            await HandleUserPromptAsync(inputText).ConfigureAwait(false);
                            return;
                        }
                        
                        // Route any other message containing text through the main handler
                        if (!string.IsNullOrWhiteSpace(inputText))
                        {
                            await HandleUserPromptAsync(inputText).ConfigureAwait(false);
                            return;
                        }
                        }
                        catch (Exception ex)
                        {
                            try { if (ShouldLogWebView2Debug()) Debug.WriteLine("[WebView2][Error] " + ex.ToString()); } catch { }
                            SendToWeb("error", null, ex.Message, addToConversation: true);
                        }
                    });

                };

                // Prefer dev URL if provided (e.g., Vite dev server)
                var devUrl = Environment.GetEnvironmentVariable("XLIFY_DEV_URL");
                if (!string.IsNullOrWhiteSpace(devUrl))
                {
                    try { _web.CoreWebView2.Navigate(devUrl); return; } catch { }
                }

                // Otherwise, serve built static files from WebApp/dist via a virtual host mapping
                var distPath = TryResolveWebDistPath();
                if (distPath != null && Directory.Exists(distPath))
                {
                    _web.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "app.xlify",
                        distPath,
                        CoreWebView2HostResourceAccessKind.Allow);
                    _web.CoreWebView2.Navigate("https://app.xlify/index.html");
                }
                else
                {
                    // Fallback minimal inlined page
                    string html = "<!doctype html><meta charset='utf-8'><style>body{font:14px Segoe UI;margin:0;padding:16px}</style>"
                                 + "<h3>XLify</h3><p>Build the web app (Vite) into WebApp/dist, or set XLIFY_DEV_URL to your dev server.</p>";
                    _web.CoreWebView2.NavigateToString(html);
                }
            }
            catch (Exception)
            {
                // WebView2 runtime likely missing; show an inline notice with link
                var panel = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
                var msg = new Label
                {
                    AutoSize = true,
                    MaximumSize = new Size(this.Width - 24, 0),
                    Text = "WebView2 runtime is not available. Install the Evergreen WebView2 Runtime to enable the embedded chat UI.",
                };
                var link = new LinkLabel
                {
                    Text = "Download WebView2 Runtime",
                    AutoSize = true,
                    LinkBehavior = LinkBehavior.HoverUnderline,
                    Top = msg.Bottom + 8,
                };
                link.Click += (s, e) =>
                {
                    try { System.Diagnostics.Process.Start("https://developer.microsoft.com/microsoft-edge/webview2/"); } catch { }
                };

                panel.Controls.Add(msg);
                panel.Controls.Add(link);
                msg.Location = new Point(12, 12);
                link.Location = new Point(12, 12 + msg.Height + 8);

                // Replace the web control placeholder
                _web.Parent?.Controls.Remove(_web);
                this.Controls.Add(panel);
                panel.BringToFront();
            }
        }

        private async Task EnsureSemanticKernelAsync()
        {
            if (_kernel != null)
            {
                if (!_sessionCreated)
                {
                    try { await RoslynWorkerClient.CreateSessionAsync(_sessionId).ConfigureAwait(false); _sessionCreated = true; } catch { }
                    if (_sessionCreated && !_sessionWarmedUp)
                    {
                        try { await WarmUpExecutorAsync().ConfigureAwait(false); _sessionWarmedUp = true; } catch { }
                    }
                }
                return;
            }
            _kernel = SemanticKernelFactory.CreateKernel(_sessionId);
            try
            {
                _responseAgent = _kernel.GetRequiredService<OpenAIResponseAgent>();
            }
            catch { }
            try { await RoslynWorkerClient.CreateSessionAsync(_sessionId).ConfigureAwait(false); _sessionCreated = true; } catch { }
            if (_sessionCreated && !_sessionWarmedUp)
            {
                try { await WarmUpExecutorAsync().ConfigureAwait(false); _sessionWarmedUp = true; } catch { }
            }
        }

        private ChatHistory BuildChatHistory(string userPrompt)
        {
            var history = new ChatHistory();
            history.AddSystemMessage(
                "You are XLify, an Excel assistant.\n" +
                "- ALWAYS use the execute_csharp tool for actions.\n" +
                "- For documentation tasks, prefer the doc plugin tools: generate_workbook_overview and summarize_selection.\n" +
                "- Write top-level C# script statements only. Do NOT wrap code in classes, namespaces, or a Main method.\n" +
                "- Do NOT call Marshal.GetActiveObject; use the provided Excel Application variable (e.g., dynamic app = Application).\n" +
                "- Prefer dynamic for Excel COM objects to avoid casting issues. Example: dynamic sheet = app.ActiveSheet; dynamic cell = sheet.Cells[row, col]; cell.Value2 = ...\n" +
                "- Access cell values via .Value2 on dynamic or casted ranges; never call .Value2 on a plain 'object'.\n" +
                "- Never store Excel COM objects in variables typed as 'object'. Use 'dynamic' or cast to the concrete interop type.\n" +
                "- When using sheet.Range[...] or Cells[,], either assign to 'dynamic' or explicitly cast: Excel.Range r = (Excel.Range)sheet.Range[sheet.Cells[r1,c1], sheet.Cells[r2,c2]].\n" +
                "- When using Columns/Rows off a Range, cast before member access if not using dynamic (e.g., ((Excel.Range)r.Columns[1]).ColumnWidth = 12).\n" +
                "- Optional parameters for COM calls: declare 'object missing = Type.Missing;' and pass 'missing'. NEVER declare 'Type missing' or assign Type.Missing to a System.Type variable.\n" +
                "- Do not assign arbitrary objects to System.Type. To get a runtime type, use 'var t = obj?.GetType();'. To refer to a type, use 'typeof(Excel.Range)'.\n" +
                "- Avoid System.Linq and extension methods on COM objects (e.g., do NOT call LINQ Select()). To select a range, use sheet.Range[...].Select(Type.Missing).\n" +
                "- When explicit typing is needed, use Excel.Worksheet/Excel.Range/Excel.Workbook; otherwise prefer dynamic to reduce COM interop errors.\n" +
                "- The alias 'Excel = Microsoft.Office.Interop.Excel' is pre-imported; use 'Excel.XlCalculation.*' enums (do NOT emit raw integers).\n" +
                "- Never assign numeric literals (e.g., 1, 0, -4135) to Excel enum-typed properties. Always use the enum constant, e.g., app.Calculation = Excel.XlCalculation.xlCalculationManual;\n" +
                "- When toggling calculation, use this exact pattern to capture and restore the previous enum value (not an int):\n" +
                "  Excel.XlCalculation prevCalc = (Excel.XlCalculation)app.Calculation;\n" +
                "  try { app.Calculation = Excel.XlCalculation.xlCalculationManual; /* work */ }\n" +
                "  finally { app.Calculation = prevCalc; }\n" +
                "- Use APIs that exist for the specific Excel object; branch on object type when needed.\n" +
                "- For PivotTables: first check if pvt.PivotCache().OLAP is true. If true, use pvt.CubeFields with exact captions; if false, use pvt.PivotFields with exact source header text. Always pvt.RefreshTable() after changes.\n" +
                "- When writing dates, assign .Value2 with DateTime.ToOADate() doubles.\n" +
                "- Performance: prefer batch operations to minimize COM calls. Build an object[,] array and assign it to a Range in one call instead of looping per cell.\n" +
                "- Performance: wrap large updates with app.ScreenUpdating=false, app.DisplayAlerts=false, and app.Calculation=Excel.XlCalculation.xlCalculationManual; restore original settings in a finally block using Excel.XlCalculation.xlCalculationAutomatic.\n" +
                "- Performance: avoid AutoFit, Select, or FreezePanes inside loops; if needed, do them once after writing data.\n" +
                "- Performance: prefer fixed ColumnWidth over AutoFit to avoid expensive layout passes.\n" +
                "- Performance: format only the used range (e.g., the exact Resize of output) instead of entire columns or sheets.\n" +
                "- Performance: avoid clearing entire sheets; overwrite the destination range or use ClearContents on just that range.\n" +
                "- Avoid magic numbers for Excel enums (e.g., use Excel.XlCalculation.xlCalculationManual instead of -4135).\n" +
                "- When you need to run code, call execute_csharp with ONLY the code; do not inline code in your chat reply.\n" +
                "- If a compile/runtime error occurs, read it and attempt one repair; if still failing, ask a concise clarifying question.\n" +
                "- Keep responses brief: explain changes and key results; ask clarifying questions if uncertain."
            );
            try
            {
                var ctx = CollectExcelContext();
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                var ctxJson = serializer.Serialize(ctx);
                history.AddSystemMessage("Current Excel context: " + Truncate(ctxJson, 3500));
            }
            catch { }

            if (!string.IsNullOrWhiteSpace(_summary))
            {
                history.AddSystemMessage("Conversation summary: " + Truncate(_summary, 1500));
            }

            try { history.AddSystemMessage(BuildRecentCodeSummary()); } catch { }

            history.AddUserMessage(userPrompt ?? string.Empty);
            return history;
        }

        private async Task WarmUpExecutorAsync()
        {
            try
            {
                // Prime Roslyn compile/JIT and basic Excel COM/dynamic binding with a tiny range write.
                var code = @"dynamic app = Application;
try
{
    var oldUpd = app.ScreenUpdating;
    app.ScreenUpdating = false;
    dynamic wb = app.ActiveWorkbook;
    dynamic sheet = wb.ActiveSheet;
    var tmp = new object[1,1];
    tmp[0,0] = "";
    var r = sheet.Range[""A1""].Resize[1,1];
    r.Value2 = tmp;
    r.ClearContents();
    app.ScreenUpdating = oldUpd;
}
catch { }";
                await RoslynWorkerClient.ExecuteAsync(code, _sessionId, timeoutMs: 4000).ConfigureAwait(false);
            }
            catch { }
        }

        private static string ExtractChatText(Microsoft.SemanticKernel.ChatMessageContent message)
        {
            if (message == null) return null;
            if (!string.IsNullOrWhiteSpace(message.Content)) return message.Content;
            try
            {
                var sb = new StringBuilder();
                if (message.Items != null)
                {
                    foreach (var item in message.Items)
                    {
                        if (item is TextContent tc && !string.IsNullOrWhiteSpace(tc.Text))
                        {
                            sb.Append(tc.Text);
                        }
                        else
                        {
                            // Try to pull a Result property via reflection to capture tool output when available
                            try
                            {
                                var prop = item.GetType().GetProperty("Result");
                                var val = prop?.GetValue(item);
                                if (val != null)
                                {
                                    sb.Append(val.ToString());
                                    continue;
                                }
                            }
                            catch { }
                            sb.Append(item.ToString());
                        }
                    }
                }
                var text = sb.ToString();
                if (!string.IsNullOrWhiteSpace(text)) return text;
            }
            catch { }
            return message.ToString();
        }

        private void LogModelToolCalls(Microsoft.SemanticKernel.ChatMessageContent message)
        {
            if (message == null || message.Items == null) return;
            foreach (var item in message.Items)
            {
                if (item == null) continue;
                var typeName = item.GetType().Name;

                // Function/Tool call from the model
                if (typeName.EndsWith("FunctionCallContent", StringComparison.OrdinalIgnoreCase) ||
                    typeName.EndsWith("ToolCallContent", StringComparison.OrdinalIgnoreCase))
                {
                    string id = GetPropString(item, "Id") ?? GetPropString(item, "CallId");
                    string name = GetPropString(item, "Name") ?? GetPropString(item, "FunctionName");
                    string args = SerializeArgs(GetPropObject(item, "Arguments"))
                                   ?? GetPropString(item, "Json")
                                   ?? item.ToString();
                    Log.Information("SK-MSG TOOL-CALL: name={Name} id={Id} args={Args}", name, id, Truncate(args, 2000));
                    try { UpdateStatus("Working…"); } catch { }
                    continue;
                }

                // Function/Tool result returned to the model
                if (typeName.EndsWith("FunctionResultContent", StringComparison.OrdinalIgnoreCase) ||
                    typeName.EndsWith("ToolResultContent", StringComparison.OrdinalIgnoreCase))
                {
                    string id = GetPropString(item, "Id") ?? GetPropString(item, "CallId");
                    string name = GetPropString(item, "Name") ?? GetPropString(item, "FunctionName");
                    string result = GetPropString(item, "Result") ?? GetPropString(item, "Content") ?? item.ToString();
                    Log.Information("SK-MSG TOOL-RESULT: name={Name} id={Id} result={Result}", name, id, Truncate(result, 2000));
                    try { UpdateStatus("Thinking…"); } catch { }
                    continue;
                }
            }
        }

        private static string GetPropString(object obj, string prop)
        {
            try
            {
                if (obj == null) return null;
                var p = obj.GetType().GetProperty(prop);
                if (p == null) return null;
                var v = p.GetValue(obj);
                return v?.ToString();
            }
            catch { return null; }
        }

        private static object GetPropObject(object obj, string prop)
        {
            try
            {
                if (obj == null) return null;
                var p = obj.GetType().GetProperty(prop);
                if (p == null) return null;
                return p.GetValue(obj);
            }
            catch { return null; }
        }

        private static string SerializeArgs(object args)
        {
            try
            {
                if (args == null) return null;
                // If it's already a string (JSON), return it
                if (args is string s) return s;
                // If it is an IEnumerable of key/value pairs, format them
                if (args is System.Collections.IEnumerable enumerable)
                {
                    var parts = new System.Collections.Generic.List<string>();
                    foreach (var entry in enumerable)
                    {
                        if (entry == null) continue;
                        var t = entry.GetType();
                        var pk = t.GetProperty("Key");
                        var pv = t.GetProperty("Value");
                        if (pk != null && pv != null)
                        {
                            var key = pk.GetValue(entry)?.ToString();
                            var val = pv.GetValue(entry);
                            if (string.Equals(key, "code", StringComparison.OrdinalIgnoreCase) && val != null)
                            {
                                var codeStr = val.ToString() ?? string.Empty;
                                var len = codeStr.Length;
                                var sha = ComputeSha1Hex(codeStr);
                                var logCode = Environment.GetEnvironmentVariable("XLIFY_LOG_CODE");
                                if (!string.IsNullOrEmpty(logCode) && (logCode.Equals("1") || logCode.Equals("true", StringComparison.OrdinalIgnoreCase)))
                                {
                                    parts.Add($"code=[{len} chars] sha1={sha} content={Truncate(codeStr, 500)}");
                                }
                                else
                                {
                                    parts.Add($"code=[{len} chars] sha1={sha} (hidden)");
                                }
                            }
                            else
                            {
                                parts.Add($"{key}={val}");
                            }
                        }
                        else
                        {
                            parts.Add(entry.ToString());
                        }
                    }
                    if (parts.Count > 0) return string.Join(", ", parts);
                }
                // Fallback to ToString
                return args.ToString();
            }
            catch { return null; }
        }

        private static string ComputeSha1Hex(string input)
        {
            try
            {
                using (var sha1 = System.Security.Cryptography.SHA1.Create())
                {
                    var bytes = System.Text.Encoding.UTF8.GetBytes(input ?? string.Empty);
                    var hash = sha1.ComputeHash(bytes);
                    var sb = new System.Text.StringBuilder(hash.Length * 2);
                    foreach (var b in hash) sb.Append(b.ToString("x2"));
                    return sb.ToString();
                }
            }
            catch { return string.Empty; }
        }

        private async Task HandleUserPromptAsync(string inputText)
        {
            AppendConversation("user", inputText);

            if (!ApiKeyVault.Has())
            {
                SendToWeb("error", null, "Add your OpenAI API key to continue.", addToConversation: true);
                return;
            }

            // No HTTP fallback; use Semantic Kernel Agents with OpenAI Responses exclusively.

            await _chatLock.WaitAsync().ConfigureAwait(false);
            try
            {
                await EnsureSemanticKernelAsync().ConfigureAwait(false);

                var history = BuildChatHistory(inputText);

                // Web search prefetch removed; rely on search-enabled Responses model via agent

                if (_responseAgent == null)
                {
                    throw new InvalidOperationException("OpenAI Responses agent is not available.");
                }

                Microsoft.SemanticKernel.ChatMessageContent response = null;
                var messages = new System.Collections.Generic.List<Microsoft.SemanticKernel.ChatMessageContent>();
                foreach (var m in history)
                {
                    messages.Add(m);
                }
                // Enable auto tool invocation for the agent via prompt execution settings
                var openAiExec = new OpenAIPromptExecutionSettings
                {
                    ToolCallBehavior = ToolCallBehavior.AutoInvokeKernelFunctions
                };
                var agentArgs = new KernelArguments(openAiExec);

                // Prefer explicit OpenAI Responses invocation options when available
                Microsoft.SemanticKernel.Agents.AgentInvokeOptions invokeOptions = null;
                try
                {
                    var t = typeof(OpenAIResponseAgent).Assembly.GetType("Microsoft.SemanticKernel.Agents.OpenAI.OpenAIResponseAgentInvokeOptions");
                    if (t != null)
                    {
                        var tmp = Activator.CreateInstance(t);
                        // Assign KernelArguments via the common property name 'Arguments'
                        var pArgs = t.GetProperty("Arguments");
                        if (pArgs != null) pArgs.SetValue(tmp, agentArgs, null);

                        // Best-effort: set ResponseCreationOptions.Tools to enable web_search when supported by SDK
                        try
                        {
                            var pResp = t.GetProperty("ResponseCreationOptions");
                            if (pResp != null)
                            {
                                var respOptionsType = pResp.PropertyType;
                                var respOptions = Activator.CreateInstance(respOptionsType);
                                var toolsProp = respOptionsType.GetProperty("Tools");
                                if (toolsProp != null && toolsProp.CanWrite)
                                {
                                    // Try to set a single tool entry with type = "web_search" via dynamic model factory if available
                                    // If the SDK surface changes, this silently no-ops.
                                    toolsProp.SetValue(respOptions, new object[] { }, null);
                                }
                                pResp.SetValue(tmp, respOptions, null);
                            }
                        }
                        catch { }

                        invokeOptions = (Microsoft.SemanticKernel.Agents.AgentInvokeOptions)tmp;
                    }
                }
                catch { }
                if (invokeOptions == null)
                {
                    var basic = new Microsoft.SemanticKernel.Agents.AgentInvokeOptions();
                    try
                    {
                        var tOpt = basic.GetType();
                        var pArgs = tOpt.GetProperty("Arguments") ?? tOpt.GetProperty("KernelArguments");
                        if (pArgs != null && pArgs.CanWrite)
                        {
                            pArgs.SetValue(basic, agentArgs, null);
                        }
                    }
                    catch { }
                    invokeOptions = basic;
                }

                // Indicate the model is thinking while we call the API
                try { UpdateStatus("Thinking…"); } catch { }

                await foreach (var item in _responseAgent.InvokeAsync(messages, null, invokeOptions, System.Threading.CancellationToken.None))
                {
                    // Log and surface any tool/function call transitions to the UI
                    try { LogModelToolCalls(item); } catch { }
                    response = item;
                }

                // Log item types for diagnostics (single consolidated block)
                try
                {
                    if (ShouldLogSkItems())
                    {
                        var items = response?.Items;
                        if (items != null)
                        {
                            var sbItems = new StringBuilder();
                            foreach (var item in items)
                            {
                                try { sbItems.AppendLine(item?.GetType()?.Name + ": " + item); } catch { }
                            }
                            if (sbItems.Length > 0)
                            {
                                try { Debug.WriteLine("[SK Items]\n" + sbItems.ToString()); } catch { }
                            }
                        }
                    }
                }
                catch { }
                if (response == null)
                {
                    throw new InvalidOperationException("OpenAI Responses agent returned no message.");
                }

                // (duplicate block removed)
                var text = ExtractChatText(response);
                if (string.IsNullOrWhiteSpace(text))
                {
                    text = "No response from the model.";
                }
                SendToWeb("assistant", null, text, addToConversation: true);
                try { UpdateStatus(""); } catch { }
            }
            catch (Exception ex)
            {
                var msg = ex.Message ?? ex.ToString();
                SendToWeb("error", null, "Semantic Kernel error: " + msg, addToConversation: true);
            }
            finally
            {
                try { _chatLock.Release(); } catch { }
            }
        }

        private static string JsonEscape(string s)
        {
            if (s == null) return string.Empty;
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");
        }

        private static bool ShouldLogWebView2Debug()
        {
            try
            {
                var v = Environment.GetEnvironmentVariable("XLIFY_LOG_WEBVIEW2_DEBUG");
                if (string.IsNullOrWhiteSpace(v)) return false; // default off
                return v.Equals("1") || v.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        private static bool ShouldLogSkItems()
        {
            try
            {
                var v = Environment.GetEnvironmentVariable("XLIFY_LOG_SK_ITEMS");
                if (string.IsNullOrWhiteSpace(v)) return false; // default: off to keep logs clean
                return v.Equals("1") || v.Equals("true", StringComparison.OrdinalIgnoreCase);
            }
            catch { return false; }
        }

        

        // Conversation storage (step 1): capture user/assistant messages to a rolling log on disk
        internal sealed class ConversationEntry
        {
            public string role { get; set; }
            public string text { get; set; }
            public DateTime ts { get; set; }
        }

        private static readonly System.Collections.Generic.List<ConversationEntry> _conversation = new System.Collections.Generic.List<ConversationEntry>();
        private static string _conversationPath;
        private static string _summaryPath;
        private static string _summary;

        // Executed code log (recent C# snippets and outcomes)
        internal sealed class CodeEntry
        {
            public DateTime ts { get; set; }
            public bool success { get; set; }
            public string info { get; set; }
            public string code { get; set; }
        }

        private static readonly System.Collections.Generic.List<CodeEntry> _codeLog = new System.Collections.Generic.List<CodeEntry>();
        private static string _codePath;

        private static void EnsureConversationPath()
        {
            if (!string.IsNullOrEmpty(_conversationPath)) return;
            try
            {
                var root = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "XLify", "Conversations");
                System.IO.Directory.CreateDirectory(root);
                _conversationPath = System.IO.Path.Combine(root, "current.json");
                _summaryPath = System.IO.Path.Combine(root, "summary.txt");
                _codePath = System.IO.Path.Combine(root, "code.json");
            }
            catch { _conversationPath = null; }
        }

        private static void AppendConversation(string role, string text)
        {
            try
            {
                EnsureConversationPath();
                if (string.IsNullOrWhiteSpace(role)) role = "system";
                if (text == null) text = string.Empty;
                _conversation.Add(new ConversationEntry { role = role, text = text, ts = DateTime.UtcNow });
                try
                {
                    if (!string.IsNullOrEmpty(_conversationPath))
                    {
                        var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                        var json = serializer.Serialize(_conversation);
                        System.IO.File.WriteAllText(_conversationPath, json, Encoding.UTF8);
                    }
                    RebuildSummary();
                }
                catch { }
            }
            catch { }
        }

        private void SendToWeb(string type, string jsonObjectOrArray, string plainText, bool addToConversation)
        {
            try
            {
                var payload = BuildJsonSafe(type, jsonObjectOrArray, plainText);
                OnUi(() => _web?.CoreWebView2?.PostWebMessageAsString(payload));
                if (addToConversation)
                {
                    if (string.Equals(type, "assistant", StringComparison.OrdinalIgnoreCase))
                        AppendConversation("assistant", plainText ?? jsonObjectOrArray);
                    else if (string.Equals(type, "error", StringComparison.OrdinalIgnoreCase))
                        AppendConversation("system", plainText ?? jsonObjectOrArray);
                }
            }
            catch { }
        }

        private void UpdateStatus(string text)
        {
            try { SendToWeb("status", null, text ?? string.Empty, addToConversation: false); } catch { }
        }

        private void OnUi(Action action)
        {
            try
            {
                if (action == null) return;
                var ctx = _uiContext ?? SynchronizationContext.Current;
                if (ctx != null)
                {
                    ctx.Post(_ => { try { action(); } catch { } }, null);
                }
                else
                {
                    action();
                }
            }
            catch { }
        }

        private Task OnUiAsync(Func<Task> action)
        {
            if (action == null) return Task.CompletedTask;
            var tcs = new TaskCompletionSource<bool>();
            var ctx = _uiContext ?? SynchronizationContext.Current;
            if (ctx != null)
            {
                try
                {
                    ctx.Post(async _ =>
                    {
                        try { await action().ConfigureAwait(false); tcs.SetResult(true); }
                        catch (Exception ex) { tcs.SetException(ex); }
                    }, null);
                }
                catch (Exception ex) { tcs.SetException(ex); }
                return tcs.Task;
            }
            // No context; run inline
            return action();
        }

        private static void RebuildSummary()
        {
            try
            {
                // Build a concise rolling summary from the most recent exchanges
                const int maxItems = 20; // last N messages
                const int maxChars = 3000; // cap summary
                var sb = new StringBuilder();
                int count = 0;
                // iterate from the end
                for (int i = _conversation.Count - 1; i >= 0 && count < maxItems; i--)
                {
                    var e = _conversation[i];
                    if (e == null) continue;
                    // Only summarize user and assistant messages (skip system/debug)
                    var role = e.role;
                    if (!string.Equals(role, "user", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(role, "assistant", StringComparison.OrdinalIgnoreCase))
                        continue;
                    var text = e.text ?? string.Empty;
                    // Collapse whitespace and shorten long lines
                    var compact = text.Replace('\r', ' ').Replace('\n', ' ');
                    if (compact.Length > 400) compact = compact.Substring(0, 400) + "...";
                    sb.Insert(0, (string.Equals(role, "user", StringComparison.OrdinalIgnoreCase) ? "User: " : "Assistant: ") + compact + "\n");
                    count++;
                }
                var summary = sb.ToString();
                if (summary.Length > maxChars) summary = summary.Substring(summary.Length - maxChars);
                _summary = summary;
                if (!string.IsNullOrEmpty(_summaryPath))
                {
                    System.IO.File.WriteAllText(_summaryPath, _summary ?? string.Empty, Encoding.UTF8);
                }
            }
            catch { }
        }

        private static string BuildContextDigest(object ctx)
        {
            try
            {
                object v;
                string sheet = AsString(TryGetPath(ctx, out v, "activeSheet", "name") ? v : null);
                string sel = AsString(TryGetPath(ctx, out v, "selection", "address") ? v : null);
                int usedR = AsInt(TryGetPath(ctx, out v, "activeSheet", "used", "rows") ? v : null);
                int usedC = AsInt(TryGetPath(ctx, out v, "activeSheet", "used", "cols") ? v : null);
                int ldr = AsInt(TryGetPath(ctx, out v, "activeSheet", "lastDataRow") ? v : null);
                int ldc = AsInt(TryGetPath(ctx, out v, "activeSheet", "lastDataCol") ? v : null);
                int tables = CountSeq(TryGetPath(ctx, out v, "tables") ? v : null);
                int pivots = CountSeq(TryGetPath(ctx, out v, "pivots") ? v : null);
                int charts = CountSeq(TryGetPath(ctx, out v, "charts") ? v : null);
                var sb = new StringBuilder();
                if (!string.IsNullOrEmpty(sheet)) sb.Append("sheet=").Append(sheet).Append(' ');
                if (!string.IsNullOrEmpty(sel)) sb.Append("sel=").Append(sel).Append(' ');
                if (usedR > 0 || usedC > 0) sb.Append("used=").Append(usedR).Append('x').Append(usedC).Append(' ');
                if (ldr > 0 || ldc > 0) sb.Append("last=").Append(ldr).Append('x').Append(ldc).Append(' ');
                if (tables >= 0) sb.Append("tables=").Append(tables).Append(' ');
                if (pivots >= 0) sb.Append("pivots=").Append(pivots).Append(' ');
                if (charts >= 0) sb.Append("charts=").Append(charts).Append(' ');
                var s = sb.ToString().Trim();
                return string.IsNullOrEmpty(s) ? "(no digest)" : s;
            }
            catch { return "(no digest)"; }
        }

        private static string BuildRunInfo(string assistantTextOrError, string stdoutText, string stderrText, string digestBefore, string digestAfter)
        {
            try
            {
                string outPart = string.IsNullOrWhiteSpace(stdoutText) ? null : Truncate(stdoutText, 300);
                string errPart = string.IsNullOrWhiteSpace(stderrText) ? null : Truncate(stderrText, 200);
                var sb = new StringBuilder();
                if (!string.IsNullOrWhiteSpace(assistantTextOrError)) sb.Append(Truncate(assistantTextOrError, 200));
                if (!string.IsNullOrWhiteSpace(outPart)) sb.Append(" | out: ").Append(outPart);
                if (!string.IsNullOrWhiteSpace(errPart)) sb.Append(" | err: ").Append(errPart);
                if (!string.IsNullOrWhiteSpace(digestBefore) || !string.IsNullOrWhiteSpace(digestAfter))
                {
                    sb.Append(" | ctx: ");
                    if (!string.IsNullOrWhiteSpace(digestBefore)) sb.Append("before[").Append(Truncate(digestBefore, 200)).Append("]");
                    if (!string.IsNullOrWhiteSpace(digestAfter)) sb.Append(" after[").Append(Truncate(digestAfter, 200)).Append("]");
                }
                return sb.ToString();
            }
            catch { return assistantTextOrError ?? string.Empty; }
        }

        private static string Truncate(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Length > max ? s.Substring(0, max) + "..." : s;
        }

        private static bool TryGetPath(object root, out object value, params string[] path)
        {
            value = root;
            try
            {
                foreach (var name in path)
                {
                    if (value == null) return false;
                    var dict = value as System.Collections.IDictionary;
                    if (dict != null)
                    {
                        if (!dict.Contains(name)) return false;
                        value = dict[name];
                        continue;
                    }
                    var t = value.GetType();
                    var prop = t.GetProperty(name);
                    if (prop == null) return false;
                    value = prop.GetValue(value, null);
                }
                return true;
            }
            catch { value = null; return false; }
        }

        private static int CountSeq(object seq)
        {
            try
            {
                if (seq == null) return -1;
                if (seq is System.Array a) return a.Length;
                if (seq is System.Collections.ICollection c) return c.Count;
                int n = 0; foreach (var _ in (System.Collections.IEnumerable)seq) n++; return n;
            }
            catch { return -1; }
        }

        private static int AsInt(object v)
        {
            try { if (v == null) return 0; return Convert.ToInt32(v); } catch { return 0; }
        }
        private static string AsString(object v)
        {
            try { return v == null ? null : v.ToString(); } catch { return null; }
        }

        internal static void AppendCodeRun(string code, bool success, string info)
        {
            try
            {
                EnsureConversationPath();
                if (code == null) code = string.Empty;
                if (info == null) info = string.Empty;
                // trim very large code to keep log lightweight
                var c = code.Length > 4000 ? code.Substring(0, 4000) + "..." : code;
                var i = info.Length > 400 ? info.Substring(0, 400) + "..." : info;
                _codeLog.Add(new CodeEntry { ts = DateTime.UtcNow, success = success, info = i, code = c });
                // persist best-effort last 20
                try
                {
                    if (!string.IsNullOrEmpty(_codePath))
                    {
                        var take = Math.Min(_codeLog.Count, 20);
                        var slice = _codeLog.GetRange(Math.Max(0, _codeLog.Count - take), take);
                        var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                        var json = serializer.Serialize(slice);
                        System.IO.File.WriteAllText(_codePath, json, Encoding.UTF8);
                    }
                }
                catch { }
            }
            catch { }
        }

        private static string BuildRecentCodeSummary()
        {
            try
            {
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                // prepare a compact array of last few runs
                var list = new System.Collections.Generic.List<object>();
                int added = 0;
                for (int i = _codeLog.Count - 1; i >= 0 && added < 5; i--)
                {
                    var e = _codeLog[i];
                    if (e == null) continue;
                    list.Add(new { ts = e.ts, success = e.success, info = e.info, code = e.code });
                    added++;
                }
                return list.Count > 0 ? ("Recent code: " + serializer.Serialize(list.ToArray())) : "Recent code: (none)";
            }
            catch { return "Recent code: (none)"; }
        }

        // Removed: IsRateLimitMessage (unused)

        private static string BuildJsonSafe(string type, string jsonObjectOrArray, string plainText)
        {
            // Use JavaScriptSerializer to ensure valid JSON for PostWebMessageAsJson
            var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
            try
            {
                if (!string.IsNullOrEmpty(jsonObjectOrArray))
                {
                    var trimmed = jsonObjectOrArray.Trim();
                    if (trimmed.StartsWith("{") || trimmed.StartsWith("["))
                    {
                        try
                        {
                            var obj = serializer.DeserializeObject(trimmed);
                            return serializer.Serialize(new { type = type, data = obj });
                        }
                        catch
                        {
                            // fallthrough to text
                        }
                    }
                }
            }
            catch { }
            return serializer.Serialize(new { type = type, text = plainText ?? string.Empty });
        }

        // Removed: legacy HTTP payload helpers (ExtractAssistantContent, TryParseAssistantJson, BuildChatMessages)

        private static object CollectExcelContext()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null) return new { error = "No application" };

                var wb = app.ActiveWorkbook;
                var ws = app.ActiveSheet as Excel.Worksheet;
                object selectionObj = null; try { selectionObj = app.Selection; } catch { }
                var selection = selectionObj as Excel.Range;
                string workbookName = null, workbookPath = null, activeSheetName = null, selectionAddress = null;
                int sheets = 0, usedRows = 0, usedCols = 0, selRows = 0, selCols = 0;
                object[] headers = null;
                object[] pivotTables = null;
                object[] charts = null;
                string selectionPivotName = null, selectionPivotAddress = null;
                string excelVersion = null;
                string selectionType = null;
                object[] sheetNames = null;
                object[] tables = null;
                object[] names = null;
                object[][] selectionSample = null;
                object selectionSummary = null;
                string activeChartName = null; int activeChartType = 0; bool activeChartIsPivot = false;
                bool? workbookSaved = null, workbookReadOnly = null;
                bool? wbProtectStructure = null, wbProtectWindows = null;
                bool? wsProtectContents = null, wsProtectDrawingObjects = null, wsProtectScenarios = null;
                int? wsAllowEditRangesCount = null;
                string calcMode = null, calcState = null;
                bool? useSystemSeparators = null; string decSep = null, thouSep = null, listSep = null;
                int? intlDateOrder = null, intlMeasurementSystem = null;
                int lastDataRow = 0, lastDataCol = 0;
                bool? undoRecordAvailable = null;
                // Add-ins and Power Query/Connections context
                object[] addIns = null; object[] comAddIns = null; bool hasSolver = false; bool hasAtp = false; bool hasPowerQuery = false; int? queryCount = null; int? connectionCount = null;
                // Clipboard context
                bool clipAvailable = false;
                string[] clipFormats = null;
                string clipText = null, clipHtml = null, clipHtmlFragment = null, clipCsvSample = null;
                int? clipRtfLength = null;
                int? clipImageWidth = null, clipImageHeight = null;
                string[] clipFiles = null;

                try { workbookName = wb?.Name; } catch { }
                try { workbookPath = wb?.FullName; } catch { }
                try { sheets = wb?.Worksheets?.Count ?? 0; } catch { }
                try { activeSheetName = ws?.Name; } catch { }
                try { excelVersion = app?.Version; } catch { }
                try { selectionType = selectionObj?.GetType()?.Name; } catch { }
                try { var ur = app?.GetType().InvokeMember("UndoRecord", System.Reflection.BindingFlags.GetProperty, null, app, null); undoRecordAvailable = (ur != null); } catch { }
                try { workbookSaved = wb?.Saved; } catch { }
                try { workbookReadOnly = wb?.ReadOnly; } catch { }
                try { wbProtectStructure = wb?.ProtectStructure; } catch { }
                try { wbProtectWindows = wb?.ProtectWindows; } catch { }
                try { wsProtectContents = ws?.ProtectContents; } catch { }
                try { wsProtectDrawingObjects = ws?.ProtectDrawingObjects; } catch { }
                try { wsProtectScenarios = ws?.ProtectScenarios; } catch { }
                try { calcMode = (app?.Calculation)?.ToString(); } catch { }
                try { calcState = (app?.CalculationState)?.ToString(); } catch { }
                try { useSystemSeparators = app?.UseSystemSeparators; } catch { }
                try { decSep = app?.DecimalSeparator; } catch { }
                try { thouSep = app?.ThousandsSeparator; } catch { }
                try { object v = app?.International[Excel.XlApplicationInternational.xlListSeparator]; if (v != null) listSep = v.ToString(); } catch { }
                try { object v = app?.International[Excel.XlApplicationInternational.xlDateOrder]; if (v != null) intlDateOrder = Convert.ToInt32(v); } catch { }
                try { object v = app?.International[Excel.XlApplicationInternational.xlMetric]; if (v != null) intlMeasurementSystem = Convert.ToInt32(v); } catch { }

                // Enumerate AddIns (safe wrappers)
                try
                {
                    var list = new System.Collections.Generic.List<object>();
                    foreach (Excel.AddIn ai in app?.AddIns)
                    {
                        try
                        {
                            string n = null, t = null, p = null; bool inst = false;
                            try { n = ai.Name; } catch { }
                            try { t = ai.Title; } catch { }
                            try { p = ai.FullName; } catch { }
                            try { inst = ai.Installed; } catch { }
                            list.Add(new { name = n, title = t, path = p, installed = inst });
                            var id = (n ?? "") + " " + (t ?? "") + " " + (p ?? "");
                            var idU = id.ToUpperInvariant();
                            if (!hasSolver && (idU.Contains("SOLVER.XLAM") || idU.Contains("SOLVER ADD-IN"))) hasSolver = inst || hasSolver;
                            if (!hasAtp && (idU.Contains("ATPVBAEN.XLAM") || idU.Contains("ANALYSIS TOOLPAK"))) hasAtp = inst || hasAtp;
                        }
                        catch { }
                    }
                    addIns = list.ToArray();
                }
                catch { }

                // Enumerate COM Add-ins (late-bound)
                try
                {
                    var list = new System.Collections.Generic.List<object>();
                    var coms = app?.GetType().InvokeMember("COMAddIns", System.Reflection.BindingFlags.GetProperty, null, app, null) as System.Collections.IEnumerable;
                    if (coms != null)
                    {
                        foreach (var cai in coms)
                        {
                            try
                            {
                                var t = cai.GetType();
                                string progId = null, desc = null; bool? connected = null;
                                try { progId = t.InvokeMember("ProgId", System.Reflection.BindingFlags.GetProperty, null, cai, null) as string; } catch { }
                                try { desc = t.InvokeMember("Description", System.Reflection.BindingFlags.GetProperty, null, cai, null) as string; } catch { }
                                try { var c = t.InvokeMember("Connect", System.Reflection.BindingFlags.GetProperty, null, cai, null); if (c != null) connected = Convert.ToBoolean(c); } catch { }
                                list.Add(new { progId = progId, description = desc, connected = connected });
                            }
                            catch { }
                        }
                    }
                    comAddIns = list.ToArray();
                }
                catch { }

                // Power Query / Queries and Connections
                try { connectionCount = wb?.Connections?.Count; } catch { }
                try
                {
                    var qColl = wb?.GetType()?.InvokeMember("Queries", System.Reflection.BindingFlags.GetProperty, null, wb, null) as System.Collections.IEnumerable;
                    if (qColl != null)
                    {
                        int cnt = 0;
                        foreach (var _ in qColl) { cnt++; }
                        queryCount = cnt;
                        hasPowerQuery = cnt > 0;
                    }
                }
                catch { }

                try
                {
                    var used = ws?.UsedRange;
                    usedRows = used?.Rows?.Count ?? 0;
                    usedCols = used?.Columns?.Count ?? 0;
                }
                catch { }

                // Last data row/column via Find
                try
                {
                    var missing = Type.Missing;
                    var lastByRows = ws?.Cells?.Find("*", ws.Cells[1, 1], missing, missing, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, missing, missing);
                    if (lastByRows != null) { try { lastDataRow = lastByRows.Row; } catch { } }
                    var lastByCols = ws?.Cells?.Find("*", ws.Cells[1, 1], missing, missing, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious, false, missing, missing);
                    if (lastByCols != null) { try { lastDataCol = lastByCols.Column; } catch { } }
                }
                catch { }

                try
                {
                    selectionAddress = selection?.Address[false, false];
                    selRows = selection?.Rows?.Count ?? 0;
                    selCols = selection?.Columns?.Count ?? 0;
                }
                catch { }

                // Selection pivot context (if selection is within a PivotTable)
                try
                {
                    var ptSel = selection?.PivotTable as Excel.PivotTable;
                    if (ptSel != null)
                    {
                        selectionPivotName = ptSel.Name;
                        try { selectionPivotAddress = ptSel.TableRange2?.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing); } catch { }
                    }
                }
                catch { }

                // Enumerate pivot tables on active sheet
                try
                {
                    var list = new System.Collections.Generic.List<object>();
                    var pts = ws?.PivotTables() as Excel.PivotTables;
                    if (pts != null)
                    {
                        int count = 0; try { count = pts.Count; } catch { }
                        for (int i = 1; i <= count; i++)
                        {
                            try
                            {
                                var pt = pts.Item(i) as Excel.PivotTable;
                                if (pt == null) continue;
                                string name = null, addr = null;
                                try { name = pt.Name; } catch { }
                                try { addr = pt.TableRange2?.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing); } catch { }
                                var rows = new System.Collections.Generic.List<string>();
                                var cols = new System.Collections.Generic.List<string>();
                                var pages = new System.Collections.Generic.List<string>();
                                var datas = new System.Collections.Generic.List<string>();
                                try { var rf = pt.RowFields() as Excel.PivotFields; int rc = rf?.Count ?? 0; for (int r = 1; r <= rc; r++) { try { rows.Add((rf.Item(r) as Excel.PivotField)?.Name); } catch { } } } catch { }
                                try { var cf = pt.ColumnFields() as Excel.PivotFields; int cc = cf?.Count ?? 0; for (int c = 1; c <= cc; c++) { try { cols.Add((cf.Item(c) as Excel.PivotField)?.Name); } catch { } } } catch { }
                                try { var pf = pt.PageFields() as Excel.PivotFields; int pc = pf?.Count ?? 0; for (int p = 1; p <= pc; p++) { try { pages.Add((pf.Item(p) as Excel.PivotField)?.Name); } catch { } } } catch { }
                                try { var df = pt.DataFields as Excel.PivotFields; int dc = df?.Count ?? 0; for (int d = 1; d <= dc; d++) { try { datas.Add((df.Item(d) as Excel.PivotField)?.Name); } catch { } } } catch { }
                                list.Add(new { name = name, address = addr, rows = rows.ToArray(), cols = cols.ToArray(), pages = pages.ToArray(), data = datas.ToArray() });
                            }
                            catch { }
                        }
                    }
                    pivotTables = list.ToArray();
                }
                catch { }

                // Enumerate charts on active sheet
                try
                {
                    var list = new System.Collections.Generic.List<object>();
                    var cos = ws?.ChartObjects() as Excel.ChartObjects;
                    if (cos != null)
                    {
                        int count = 0; try { count = cos.Count; } catch { }
                        for (int i = 1; i <= count; i++)
                        {
                            try
                            {
                                var co = cos.Item(i) as Excel.ChartObject;
                                var ch = co?.Chart;
                                string name = null; try { name = co?.Name; } catch { }
                                int chartType = 0; try { chartType = (int)ch.ChartType; } catch { }
                                bool isPivot = false; try { var pl = ch.PivotLayout; isPivot = pl != null; } catch { isPivot = false; }
                                list.Add(new { name = name, type = chartType, isPivot = isPivot });
                            }
                            catch { }
                        }
                    }
                    charts = list.ToArray();
                }
                catch { }

                // Active chart context (if any)
                try
                {
                    var ach = app.ActiveChart;
                    if (ach != null)
                    {
                        try { activeChartName = ach.Name; } catch { }
                        try { activeChartType = (int)ach.ChartType; } catch { }
                        try { var pl = ach.PivotLayout; activeChartIsPivot = pl != null; } catch { activeChartIsPivot = false; }
                    }
                }
                catch { }

                // Worksheets list
                try
                {
                    var list = new System.Collections.Generic.List<string>();
                    var wss = wb?.Worksheets as Excel.Sheets;
                    int count = 0; try { count = wss?.Count ?? 0; } catch { }
                    for (int i = 1; i <= count; i++)
                    {
                        try { var wsi = wss[i] as Excel.Worksheet; if (wsi != null) list.Add(wsi.Name); } catch { }
                    }
                    sheetNames = list.ToArray();
                }
                catch { }

                // Tables (ListObjects) on active sheet
                try
                {
                    var list = new System.Collections.Generic.List<object>();
                    var los = ws?.ListObjects as Excel.ListObjects;
                    int count = 0; try { count = los?.Count ?? 0; } catch { }
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            var lo = los[i] as Excel.ListObject;
                            if (lo == null) continue;
                            string name = null, addr = null; object[] loHeaders = null;
                            try { name = lo.Name; } catch { }
                            try { addr = lo.Range?.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing); } catch { }
                            try
                            {
                                var hr = lo.HeaderRowRange as Excel.Range;
                                if (hr != null)
                                {
                                    object[,] data = hr.Value2 as object[,];
                                    if (data != null)
                                    {
                                        int cols = data.GetLength(1);
                                        var h = new object[cols];
                                        for (int c = 1; c <= cols; c++) h[c - 1] = data[1, c];
                                        loHeaders = h;
                                    }
                                }
                            }
                            catch { }
                            list.Add(new { name = name, address = addr, headers = loHeaders });
                        }
                        catch { }
                    }
                    tables = list.ToArray();
                }
                catch { }

                // Workbook named ranges
                try
                {
                    var list = new System.Collections.Generic.List<object>();
                    var nms = wb?.Names as Excel.Names;
                    int count = 0; try { count = nms?.Count ?? 0; } catch { }
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            var nm = nms.Item(i) as Excel.Name;
                            if (nm == null) continue;
                            string name = null, addr = null, refersTo = null;
                            try { name = nm.Name; } catch { }
                            try { refersTo = nm.RefersTo; } catch { }
                            try { addr = nm.RefersToRange?.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing); } catch { }
                            list.Add(new { name = name, address = addr, refersTo = refersTo });
                        }
                        catch { }
                    }
                    names = list.ToArray();
                }
                catch { }

                // Selection sample values (up to 10x10)
                try
                {
                    if (selection != null)
                    {
                        int r1 = 0, c1 = 0;
                        try { r1 = selection.Row; c1 = selection.Column; } catch { }
                        int rows = Math.Min(selection?.Rows?.Count ?? 0, 10);
                        int cols = Math.Min(selection?.Columns?.Count ?? 0, 10);
                        if (rows > 0 && cols > 0)
                        {
                            var topLeft = selection.Worksheet.Cells[r1, c1];
                            var bottomRight = selection.Worksheet.Cells[r1 + rows - 1, c1 + cols - 1];
                            var sampleRange = selection.Worksheet.Range[topLeft, bottomRight] as Excel.Range;
                            object value = null; try { value = sampleRange?.Value2; } catch { }
                            var result = new System.Collections.Generic.List<object[]>();
                            int selNum = 0, selTxt = 0, selBlank = 0, selBool = 0, selOther = 0;
                            if (value is object[, ] data)
                            {
                                for (int r = 1; r <= data.GetLength(0); r++)
                                {
                                    var row = new object[cols];
                                    for (int c = 1; c <= cols; c++)
                                    {
                                        var cell = data[r, c];
                                        row[c - 1] = cell;
                                        if (cell == null) selBlank++;
                                        else if (cell is bool) selBool++;
                                        else if (cell is string) selTxt++;
                                        else if (cell is sbyte || cell is byte || cell is short || cell is ushort || cell is int || cell is uint || cell is long || cell is ulong || cell is float || cell is double || cell is decimal)
                                            selNum++;
                                        else selOther++;
                                    }
                                    result.Add(row);
                                }
                            }
                            else if (value != null)
                            {
                                result.Add(new object[] { value });
                                if (value is bool) selBool++; else if (value is string) selTxt++; else if (value is sbyte || value is byte || value is short || value is ushort || value is int || value is uint || value is long || value is ulong || value is float || value is double || value is decimal) selNum++; else selOther++;
                            }
                            else { selBlank++; }
                            selectionSample = result.ToArray();
                            // attach summary alongside selection later via selectionSummary object
                            selectionSummary = new { numeric = selNum, text = selTxt, blanks = selBlank, logical = selBool, other = selOther };
                        }
                    }
                }
                catch { }

                // Clipboard snapshot (best-effort, small sample only)
                try
                {
                    System.Windows.Forms.IDataObject data = null;
                    for (int i = 0; i < 3; i++)
                    {
                        try { data = System.Windows.Forms.Clipboard.GetDataObject(); break; }
                        catch { try { System.Threading.Thread.Sleep(20); } catch { } }
                    }
                    if (data != null)
                    {
                        clipAvailable = true;
                        try { clipFormats = data.GetFormats(); } catch { }

                        try
                        {
                            if (data.GetDataPresent(System.Windows.Forms.DataFormats.UnicodeText))
                            {
                                var t = data.GetData(System.Windows.Forms.DataFormats.UnicodeText) as string;
                                if (!string.IsNullOrEmpty(t)) clipText = t.Length > 2000 ? t.Substring(0, 2000) : t;
                            }
                            else if (data.GetDataPresent(System.Windows.Forms.DataFormats.Text))
                            {
                                var t = data.GetData(System.Windows.Forms.DataFormats.Text) as string;
                                if (!string.IsNullOrEmpty(t)) clipText = t.Length > 2000 ? t.Substring(0, 2000) : t;
                            }
                        }
                        catch { }

                        try
                        {
                            if (data.GetDataPresent(System.Windows.Forms.DataFormats.Html))
                            {
                                var html = data.GetData(System.Windows.Forms.DataFormats.Html) as string;
                                if (!string.IsNullOrEmpty(html))
                                {
                                    clipHtml = html.Length > 4000 ? html.Substring(0, 4000) : html;
                                    // Attempt to extract CF_HTML fragment
                                    try
                                    {
                                        int si = html.IndexOf("StartFragment:");
                                        int ei = html.IndexOf("EndFragment:");
                                        if (si >= 0 && ei > si)
                                        {
                                            int start = 0, end = 0;
                                            int.TryParse(html.Substring(si + 14, html.IndexOf('\n', si) - (si + 14)).Trim(), out start);
                                            int.TryParse(html.Substring(ei + 12, html.IndexOf('\n', ei) - (ei + 12)).Trim(), out end);
                                            if (start >= 0 && end > start && end <= html.Length)
                                            {
                                                var frag = html.Substring(start, end - start);
                                                if (!string.IsNullOrEmpty(frag)) clipHtmlFragment = frag.Length > 4000 ? frag.Substring(0, 4000) : frag;
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        catch { }

                        try
                        {
                            if (data.GetDataPresent(System.Windows.Forms.DataFormats.CommaSeparatedValue))
                            {
                                var csvObj = data.GetData(System.Windows.Forms.DataFormats.CommaSeparatedValue);
                                string csv = null;
                                if (csvObj is string s) csv = s;
                                else if (csvObj is System.IO.Stream st)
                                {
                                    using (var sr = new System.IO.StreamReader(st)) csv = sr.ReadToEnd();
                                }
                                if (!string.IsNullOrEmpty(csv))
                                {
                                    var lines = csv.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                                    var take = Math.Min(lines.Length, 10);
                                    clipCsvSample = string.Join("\n", lines, 0, take);
                                }
                            }
                        }
                        catch { }

                        try
                        {
                            if (data.GetDataPresent(System.Windows.Forms.DataFormats.Rtf))
                            {
                                var rtf = data.GetData(System.Windows.Forms.DataFormats.Rtf) as string;
                                if (rtf != null) clipRtfLength = rtf.Length;
                            }
                        }
                        catch { }

                        try
                        {
                            if (System.Windows.Forms.Clipboard.ContainsImage())
                            {
                                var img = System.Windows.Forms.Clipboard.GetImage();
                                if (img != null) { clipImageWidth = img.Width; clipImageHeight = img.Height; }
                            }
                        }
                        catch { }

                        try
                        {
                            if (data.GetDataPresent(System.Windows.Forms.DataFormats.FileDrop))
                            {
                                var files = data.GetData(System.Windows.Forms.DataFormats.FileDrop) as string[];
                                if (files != null) clipFiles = files;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                try
                {
                    // Try to get header row from the first row of UsedRange (limit 50 columns)
                    if (ws?.UsedRange != null)
                    {
                        int colCount = Math.Min(ws.UsedRange.Columns.Count, 50);
                        var headerRange = ws.Range[ws.Cells[ws.UsedRange.Row, ws.UsedRange.Column], ws.Cells[ws.UsedRange.Row, ws.UsedRange.Column + colCount - 1]] as Excel.Range;
                        if (headerRange != null)
                        {
                            object[,] data = headerRange.Value2 as object[,];
                            if (data != null)
                            {
                                int cols = data.GetLength(1);
                                headers = new object[cols];
                                for (int c = 1; c <= cols; c++) headers[c - 1] = data[1, c];
                            }
                            else
                            {
                                // Single cell or 1D
                                var single = headerRange.Value2;
                                headers = new object[] { single };
                            }
                        }
                    }
                }
                catch { }

                // Worksheet protection details extras
                try { wsAllowEditRangesCount = ws?.Protection?.AllowEditRanges?.Count; } catch { }

                // Computed protection flags
                bool? workbookIsProtected = null; try { workbookIsProtected = (wbProtectStructure == true) || (wbProtectWindows == true); } catch { }
                bool? worksheetIsProtected = null; try { worksheetIsProtected = (wsProtectContents == true) || (wsProtectDrawingObjects == true) || (wsProtectScenarios == true); } catch { }

                // return context object
                return new
                {
                    workbook = new { name = workbookName, path = workbookPath, sheetCount = sheets, saved = workbookSaved, readOnly = workbookReadOnly, protection = new { structure = wbProtectStructure, windows = wbProtectWindows, isProtected = workbookIsProtected } },
                    activeSheet = new { name = activeSheetName, used = new { rows = usedRows, cols = usedCols }, lastDataRow = lastDataRow, lastDataCol = lastDataCol, protection = new { contents = wsProtectContents, drawingObjects = wsProtectDrawingObjects, scenarios = wsProtectScenarios, isProtected = worksheetIsProtected, allowEditRanges = wsAllowEditRangesCount } },
                    selection = new { type = selectionType, address = selectionAddress, rows = selRows, cols = selCols, sample = selectionSample, summary = selectionSummary },
                    selectionPivot = new { name = selectionPivotName, address = selectionPivotAddress },
                    headers = headers,
                    sheets = sheetNames,
                    tables = tables,
                    pivots = pivotTables,
                    charts = charts,
                    activeChart = new { name = activeChartName, type = activeChartType, isPivot = activeChartIsPivot },
                    names = names,
                    excel = new { version = excelVersion, culture = System.Globalization.CultureInfo.CurrentCulture?.Name, calculation = new { mode = calcMode, state = calcState }, separators = new { useSystem = useSystemSeparators, decimalSep = decSep, thousandsSep = thouSep, listSep = listSep }, international = new { dateOrder = intlDateOrder, measurementSystem = intlMeasurementSystem } },
                    addIns = new { addIns = addIns, comAddIns = comAddIns, hasSolver = hasSolver, hasAnalysisToolPak = hasAtp },
                    powerQuery = new { available = hasPowerQuery, queryCount = queryCount, connectionCount = connectionCount },
                    undo = new { supportsUndoRecord = undoRecordAvailable },
                    clipboard = new { available = clipAvailable, formats = clipFormats, text = clipText, html = clipHtml, htmlFragment = clipHtmlFragment, csvSample = clipCsvSample, rtfLength = clipRtfLength, image = new { width = clipImageWidth, height = clipImageHeight }, files = clipFiles }
                };
            }
            catch
            {
                return new { error = "Failed to collect context" };
            }
        }

        // Removed: CallOpenAIAsync


        // Removed: CallOpenAIResponsesAsync (legacy direct Responses HTTP call)


        // Removed: CallOpenAIWebSearchAsync (legacy manual web search prefetch)

        

        private static string TryResolveWebDistPath()
        {
            try
            {
                // Typical debug path: .../XLify/bin/Debug/ -> project root is two levels up
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var dir = new DirectoryInfo(baseDir);
                for (int i = 0; i < 4 && dir != null; i++) dir = dir.Parent; // climb up to solution root

                // Prefer project-local WebApp/dist
                string[] candidates = new[]
                {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "WebApp", "dist"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "WebApp", "dist"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebApp", "dist"),
                    dir != null ? Path.Combine(dir.FullName, "XLify", "WebApp", "dist") : null,
                };
                foreach (var p in candidates)
                {
                    if (p != null)
                    {
                        var full = Path.GetFullPath(p);
                        if (Directory.Exists(full)) return full;
                    }
                }
            }
            catch { }
            return null;
        }
    }
}
