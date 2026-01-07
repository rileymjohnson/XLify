using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Excel = Microsoft.Office.Interop.Excel;
using Serilog;

namespace XLify.CSharpExecutor
{
    [ComVisible(true)]
    [Guid("1E940616-2A64-4D25-A8C7-9F7509D96A5B")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ICSharpExecutor
    {
        bool Ping();
        void ShowMessage(string text, string title);
        void CreateSession(string sessionId, object excelApp);
        void ResetSession(string sessionId);
        void DestroySession(string sessionId);
        ExecutionResult ExecuteInSession(string sessionId, string code, int timeoutMs);
        ExecutionResult ExecuteOneOff(string code, object excelApp, int timeoutMs);
    }

    [ComVisible(true)]
    [Guid("AD07D5E9-5A35-4DB3-8F5D-9E8B2347BFB7")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("XLify.CSharpExecutor")]
    public sealed class CSharpExecutor : ICSharpExecutor
    {
        private static readonly ConcurrentDictionary<string, SessionState> Sessions = new ConcurrentDictionary<string, SessionState>(StringComparer.Ordinal);

        public bool Ping() => true;

        public void ShowMessage(string text, string title)
        {
            var caption = string.IsNullOrWhiteSpace(title) ? "XLify" : title;
            MessageBox.Show(text ?? string.Empty, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void CreateSession(string sessionId, object excelApp)
        {
            if (string.IsNullOrWhiteSpace(sessionId)) sessionId = "default";
            var state = Sessions.GetOrAdd(sessionId, _ => new SessionState());
            state.SetExcelApp(excelApp);
            try { Log.Information("[worker] CreateSession session={Session}", sessionId); } catch { }
        }

        public void ResetSession(string sessionId)
        {
            if (string.IsNullOrWhiteSpace(sessionId)) sessionId = "default";
            if (Sessions.TryGetValue(sessionId, out var state))
            {
                state.Reset();
                try { Log.Information("[worker] ResetSession session={Session}", sessionId); } catch { }
            }
        }

        public void DestroySession(string sessionId)
        {
            if (string.IsNullOrWhiteSpace(sessionId)) sessionId = "default";
            if (Sessions.TryRemove(sessionId, out var state))
            {
                state.Dispose();
                try { Log.Information("[worker] DestroySession session={Session}", sessionId); } catch { }
            }
        }

        public ExecutionResult ExecuteInSession(string sessionId, string code, int timeoutMs)
        {
            if (string.IsNullOrWhiteSpace(sessionId)) sessionId = "default";
            var state = Sessions.GetOrAdd(sessionId, _ => new SessionState());
            return ExecuteCore(state, code, timeoutMs, reuseState: true);
        }

        public ExecutionResult ExecuteOneOff(string code, object excelApp, int timeoutMs)
        {
            var state = new SessionState();
            state.SetExcelApp(excelApp);
            return ExecuteCore(state, code, timeoutMs, reuseState: false);
        }

        private ExecutionResult ExecuteCore(SessionState state, string code, int timeoutMs, bool reuseState)
        {
            code = NormalizeCode(code);
            var codeSha = HashUtil.ComputeSha1Hex(code);

            if (string.IsNullOrWhiteSpace(code))
            {
                return ExecutionResult.ErrorResult("No code provided");
            }

            var timeout = timeoutMs > 0 ? timeoutMs : 15000;
            var swOut = new StringWriter(new StringBuilder());
            var swErr = new StringWriter(new StringBuilder());
            var oldOut = Console.Out;
            var oldErr = Console.Error;
            var stopwatch = Stopwatch.StartNew();
            try { Log.Information("[worker] Execute start reuse={Reuse} codeLen={Len} codeSha={Sha} timeoutMs={Timeout}", reuseState, code?.Length ?? 0, codeSha, timeoutMs); } catch { }

            state.Gate.Wait();
            try
            {
                Console.SetOut(swOut);
                Console.SetError(swErr);

                var globals = new ScriptGlobals(state.ExcelApp);
                var options = state.Options;
                var cts = new CancellationTokenSource(timeout);

                // Ensure the Excel instance shadows the Excel.Application type name without breaking using order
                string BuildBootstrappedCode(string src, bool allowUsings)
                {
                    if (string.IsNullOrWhiteSpace(src)) return "var Application = (ApplicationInstance ?? ExcelApp);";

                    var lines = src.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
                    var sb = new StringBuilder();
                    int i = 0;
                    if (allowUsings)
                    {
                        for (; i < lines.Length; i++)
                        {
                            var t = lines[i].TrimStart();
                            if (string.IsNullOrWhiteSpace(t))
                            {
                                sb.AppendLine(lines[i]);
                                continue;
                            }
                            if (t.StartsWith("using ") || t.StartsWith("extern alias "))
                            {
                                sb.AppendLine(lines[i]);
                                continue;
                            }
                            break;
                        }
                        // Provide a stable alias for Excel interop types so scripts can use 'Excel.XlCalculation', etc.
                        sb.AppendLine("using Excel = Microsoft.Office.Interop.Excel;");
                    }
                    sb.AppendLine("var Application = (ApplicationInstance ?? ExcelApp);");

                    for (; i < lines.Length; i++)
                    {
                        var t = lines[i].TrimStart();
                        if (!allowUsings && (t.StartsWith("using ") || t.StartsWith("extern alias "))) continue;
                        sb.AppendLine(lines[i]);
                    }
                    return sb.ToString();
                }

                // If the incoming code defines a static Main method, append an invocation to execute it.
                string AppendMainInvocationIfPresent(string src, string prepared)
                {
                    try
                    {
                        var nsMatch = Regex.Match(src, @"namespace\s+([A-Za-z_][\w\.]*)", RegexOptions.Singleline);
                        string ns = nsMatch.Success ? nsMatch.Groups[1].Value : null;
                        var classMatch = Regex.Match(src, @"class\s+([A-Za-z_]\w*).*?static\s+void\s+Main\s*\(", RegexOptions.Singleline);
                        if (!classMatch.Success) return prepared;
                        var cls = classMatch.Groups[1].Value;
                        var fq = string.IsNullOrWhiteSpace(ns) ? cls : (ns + "." + cls);
                        var invokeLine = fq + ".Main();";
                        var sb = new StringBuilder(prepared);
                        sb.AppendLine();
                        sb.AppendLine(invokeLine);
                        return sb.ToString();
                    }
                    catch { return prepared; }
                }

                Task<ScriptState<object>> scriptTask;
                if (reuseState && state.ScriptState != null)
                {
                    var nextCode = BuildBootstrappedCode(code, allowUsings: false); // CSharpScript.ContinueWith disallows new using directives
                    nextCode = AppendMainInvocationIfPresent(code, nextCode);
                    scriptTask = state.ScriptState.ContinueWithAsync(nextCode, options, cancellationToken: cts.Token);
                }
                else
                {
                    var initialCode = BuildBootstrappedCode(code, allowUsings: true);
                    initialCode = AppendMainInvocationIfPresent(code, initialCode);
                    scriptTask = CSharpScript.RunAsync(initialCode, options, globals: globals, cancellationToken: cts.Token);
                }

                bool completed = scriptTask.Wait(timeout);
                if (!completed)
                {
                    cts.Cancel();
                    var to = ExecutionResult.TimeoutResult(timeout);
                    try { Log.Warning("[worker] Execute timeout codeSha={Sha} elapsedMs={Ms}", codeSha, stopwatch.ElapsedMilliseconds); } catch { }
                    return to;
                }

                var scriptState = scriptTask.Result;
                if (reuseState)
                {
                    state.ScriptState = scriptState;
                }

                stopwatch.Stop();
                var output = BuildOutput(swOut, swErr, scriptState?.ReturnValue);
                var ok = ExecutionResult.SuccessResult(output, stopwatch.ElapsedMilliseconds);
                try { Log.Information("[worker] Execute done success=true codeSha={Sha} elapsedMs={Ms}", codeSha, stopwatch.ElapsedMilliseconds); } catch { }
                return ok;
            }
            catch (CompilationErrorException cee)
            {
                stopwatch.Stop();
                var msg = string.Join(Environment.NewLine, cee.Diagnostics);
                try { Log.Error("[worker] Execute compile-error codeSha={Sha} elapsedMs={Ms}\n{Error}", codeSha, stopwatch.ElapsedMilliseconds, msg); } catch { }
                return ExecutionResult.ErrorResult(msg, stopwatch.ElapsedMilliseconds);
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                try { Log.Error("[worker] Execute error codeSha={Sha} elapsedMs={Ms}\n{Error}", codeSha, stopwatch.ElapsedMilliseconds, ex.ToString()); } catch { }
                return ExecutionResult.ErrorResult(ex.ToString(), stopwatch.ElapsedMilliseconds);
            }
            finally
            {
                try { Console.SetOut(oldOut); } catch { }
                try { Console.SetError(oldErr); } catch { }
                state.Gate.Release();
            }
        }

        private static string BuildOutput(StringWriter swOut, StringWriter swErr, object returnValue)
        {
            var sb = new StringBuilder();
            var outText = swOut.ToString();
            var errText = swErr.ToString();
            if (!string.IsNullOrWhiteSpace(outText)) sb.AppendLine(outText.TrimEnd());
            if (!string.IsNullOrWhiteSpace(errText)) sb.AppendLine("[stderr]").AppendLine(errText.TrimEnd());
            if (returnValue != null)
            {
                sb.AppendLine("[return]").AppendLine(returnValue.ToString());
            }
            return sb.ToString().TrimEnd();
        }

        private static string NormalizeCode(string code)
        {
            if (code == null) return string.Empty;

            code = code.Trim();

            // Strip simple wrappers (quotes or code fences) that may arrive from JSON/text transports
            if (code.StartsWith("\"") && code.EndsWith("\"") && code.Length >= 2)
            {
                code = code.Substring(1, code.Length - 2);
            }
            if (code.StartsWith("```"))
            {
                var idx = code.IndexOf('\n');
                if (idx >= 0) code = code.Substring(idx + 1);
                if (code.EndsWith("```")) code = code.Substring(0, code.Length - 3);
            }

            // Remove any stray fenced blocks or backticks that slipped through
            code = code.Replace("```", string.Empty);

            // Normalize common problematic unicode characters coming from LLMs
            code = code
                .Replace('\u201C', '"') // left smart quote
                .Replace('\u201D', '"') // right smart quote
                .Replace('\u201E', '"') // low double quote
                .Replace('\u2018', '\'') // left single smart quote
                .Replace('\u2019', '\'') // right single smart quote
                .Replace("\u200B", string.Empty) // zero-width space
                .Replace("\u200C", string.Empty) // zero-width non-joiner
                .Replace("\u200D", string.Empty) // zero-width joiner
                .Replace("\uFEFF", string.Empty); // BOM

            // If the code still looks JSON-escaped, unescape typical sequences
            var looksEscaped = code.Contains("\\n") || code.Contains("\\r") || code.Contains("\\t") || code.Contains("\\\"");
            if (looksEscaped)
            {
                try { code = Regex.Unescape(code); } catch { }
            }

            code = code.Replace("\\r\\n", "\n");
            code = code.Replace("\\n", "\n");
            code = code.Replace("\\r", "\n");
            code = code.Replace("\r\n", "\n");
            code = code.Replace("\r", "\n");

            // Convert C# 11 raw string literals ("""...""") to verbatim strings (@"...")
            // Many models emit raw strings which older language versions will parse as errors like "Newline in constant".
            try
            {
                code = System.Text.RegularExpressions.Regex.Replace(
                    code,
                    "(?<![$@])\"\"\"(.*?)\"\"\"",
                    m =>
                    {
                        var inner = m.Groups[1].Value;
                        // In a verbatim string, double quotes must be doubled
                        inner = inner.Replace("\"", "\"\"");
                        return "@\"" + inner + "\"";
                    },
                    System.Text.RegularExpressions.RegexOptions.Singleline);
            }
            catch { }
            return code;
        }
    }

    internal static class HashUtil
    {
        public static string ComputeSha1Hex(string input)
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
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("B7A4E9B8-7E12-4D25-92A5-5C3D2D65F9A5")]
    public sealed class ExecutionResult
    {
        public bool Success { get; set; }
        public bool TimedOut { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
        public double ElapsedMs { get; set; }

        public static ExecutionResult SuccessResult(string output, double elapsedMs) => new ExecutionResult { Success = true, Output = output ?? string.Empty, ElapsedMs = elapsedMs };
        public static ExecutionResult ErrorResult(string error, double elapsedMs = 0) => new ExecutionResult { Success = false, Error = error ?? string.Empty, ElapsedMs = elapsedMs };
        public static ExecutionResult TimeoutResult(int timeoutMs) => new ExecutionResult { Success = false, TimedOut = true, Error = $"Timed out after {timeoutMs} ms" };
    }

    internal sealed class SessionState : IDisposable
    {
        public SemaphoreSlim Gate { get; } = new SemaphoreSlim(1, 1);
            public ScriptState<object> ScriptState { get; set; }
            public ScriptOptions Options { get; }
            public Excel.Application ExcelApp { get; private set; }

        public SessionState()
        {
            Options = ScriptOptions.Default
                .AddReferences(
                    typeof(object).Assembly,
                    typeof(Enumerable).Assembly,
                    typeof(Task).Assembly,
                    typeof(MessageBox).Assembly,
                    typeof(Excel.Application).Assembly,
                    typeof(Microsoft.CSharp.RuntimeBinder.Binder).Assembly)
                .AddImports(
                    "System",
                    "System.IO",
                    "System.Linq",
                    "System.Text",
                    "System.Threading",
                    "System.Threading.Tasks",
                    "System.Collections.Generic",
                    "System.Windows.Forms",
                    "Microsoft.Office.Interop.Excel",
                    "Microsoft.CSharp");
        }

        public void SetExcelApp(object app)
        {
            ExcelApp = app as Excel.Application;
        }

        public void Reset()
        {
            ScriptState = null;
        }

        public void Dispose()
        {
            try
            {
                if (ExcelApp != null)
                {
                    Marshal.FinalReleaseComObject(ExcelApp);
                }
            }
            catch { }
            ExcelApp = null;
            try { Gate.Dispose(); } catch { }
        }
    }

    public sealed class ScriptGlobals
    {
        public ScriptGlobals(object excelApp)
        {
            ExcelApp = excelApp as Excel.Application;
            ApplicationInstance = ExcelApp;
        }

        public Excel.Application ExcelApp { get; }
        public Excel.Application ApplicationInstance { get; }
    }

    internal static class Program
    {
        private sealed class ComApplicationContext : ApplicationContext
        {
            private readonly RegistrationServices _registrar;
            private readonly int _cookie;

            public ComApplicationContext(RegistrationServices registrar, int cookie)
            {
                _registrar = registrar;
                _cookie = cookie;
            }

            protected override void Dispose(bool disposing)
            {
                base.Dispose(disposing);
                try { _registrar?.UnregisterTypeForComClients(_cookie); } catch { }
            }
        }

        [STAThread]
        private static void Main(string[] args)
        {
            // Configure Serilog to Seq for worker process
            try
            {
                var seqUrl = Environment.GetEnvironmentVariable("SEQ_URL");
                if (string.IsNullOrWhiteSpace(seqUrl)) seqUrl = "http://localhost:5341";
                var seqApiKey = Environment.GetEnvironmentVariable("SEQ_API_KEY");

                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()
                    .Enrich.FromLogContext()
                    .Enrich.WithProperty("MachineName", Environment.MachineName)
                    .Enrich.WithProperty("ProcessId", Process.GetCurrentProcess().Id)
                    .Enrich.WithProperty("ProcessName", Process.GetCurrentProcess().ProcessName)
                    .Enrich.WithProperty("Subsystem", "Worker")
                    .Enrich.WithProperty("App", "XLify.CSharpExecutor")
                    .Enrich.WithProperty("Workspace", "XLify-Worker")
                    .WriteTo.Seq(seqUrl, apiKey: string.IsNullOrWhiteSpace(seqApiKey) ? null : seqApiKey)
                    .CreateLogger();

                Log.Information("Worker starting");
            }
            catch { }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            RegistrationServices registrar = null;
            int cookie = 0;
            try
            {
                registrar = new RegistrationServices();
                cookie = registrar.RegisterTypeForComClients(typeof(CSharpExecutor), RegistrationClassContext.LocalServer, RegistrationConnectionType.MultipleUse);
                try { Log.Information("Worker COM registered"); } catch { }
            }
            catch (Exception ex)
            {
                try { Log.Error("COM server registration failed: {Error}", ex.Message); } catch { }
                MessageBox.Show("COM server registration failed: " + ex.Message, "XLify", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try { Application.Run(new ComApplicationContext(registrar, cookie)); }
            finally
            {
                try { Log.Information("Worker shutting down"); Log.CloseAndFlush(); } catch { }
            }
        }
    }
}
