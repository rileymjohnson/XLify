using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.SemanticKernel;
using Serilog;
using System.Security.Cryptography;
using System.Text;

namespace XLify
{
    /// <summary>
    /// Semantic Kernel tool that executes C# via the Roslyn worker for a specific session.
    /// The sessionId is fixed per task pane; SK callers don’t need to pass it.
    /// </summary>
    internal sealed class RoslynTool
    {
        private readonly string _sessionId;

        public RoslynTool(string sessionId)
        {
            _sessionId = sessionId ?? throw new ArgumentNullException(nameof(sessionId));
        }

        [KernelFunction("execute_csharp")]
        [Description("Execute C# code in the current session and return stdout/stderr.")]
        public async Task<string> ExecuteCSharpAsync(
            [Description("C# code to execute")] string code,
            [Description("Optional timeout in milliseconds")] int? timeoutMs = null)
        {
            try
            {
                var len = (code?.Length ?? 0);
                var sha = ComputeSha1Hex(code ?? string.Empty);
                System.Diagnostics.Debug.WriteLine("[RoslynTool] Executing C# (len=" + len + ") for session " + _sessionId);
                try { Log.Information("[RoslynTool] Code SHA={Sha} Length={Len} Session={Session}", sha, len, _sessionId); } catch { }
                // Optionally log full code if enabled
                var logCode = Environment.GetEnvironmentVariable("XLIFY_LOG_CODE");
                if (!string.IsNullOrEmpty(logCode) && (logCode.Equals("1") || logCode.Equals("true", StringComparison.OrdinalIgnoreCase)))
                {
                    System.Diagnostics.Debug.WriteLine("[RoslynTool Code]\n" + (code ?? string.Empty));
                }
                // Persist a pending entry so we can correlate failures if the worker crashes
                try { MyTaskPaneControl.AppendCodeRun(code ?? string.Empty, false, "submitted"); } catch { }
            }
            catch { }
            var opts = new ExecutionOptions { TimeoutMs = timeoutMs ?? 15000 };
            var hints = BuildExcelHints();

            var resp = await RoslynWorkerClient.ExecuteAsync(code ?? string.Empty, _sessionId, timeoutMs ?? 15000, options: opts, hints: hints).ConfigureAwait(false);

            try
            {
                System.Diagnostics.Debug.WriteLine("[RoslynTool] Success=" + resp?.Success + ", Err=" + resp?.Error + ", OutputLen=" + (resp?.Output?.Length ?? 0) + ", Hwnd=" + hints?.Hwnd + ", Pid=" + hints?.ProcessId);
            }
            catch { }

            if (resp == null)
            {
                try { MyTaskPaneControl.AppendCodeRun(code ?? string.Empty, false, "null response from worker"); } catch { }
                return "Execution failed: null response from worker";
            }

            if (!resp.Success)
            {
                if (resp.TimedOut)
                {
                    try { MyTaskPaneControl.AppendCodeRun(code ?? string.Empty, false, "timed out"); } catch { }
                    return "Execution timed out";
                }
                var err = !string.IsNullOrWhiteSpace(resp.Error) ? resp.Error : "Execution failed";
                try { MyTaskPaneControl.AppendCodeRun(code ?? string.Empty, false, err); } catch { }
                return err;
            }

            try { MyTaskPaneControl.AppendCodeRun(code ?? string.Empty, true, resp.Output ?? string.Empty); } catch { }
            return resp.Output ?? string.Empty;
        }

        private static string ComputeSha1Hex(string input)
        {
            try
            {
                using (var sha1 = SHA1.Create())
                {
                    var bytes = Encoding.UTF8.GetBytes(input ?? string.Empty);
                    var hash = sha1.ComputeHash(bytes);
                    var sb = new StringBuilder(hash.Length * 2);
                    foreach (var b in hash) sb.Append(b.ToString("x2"));
                    return sb.ToString();
                }
            }
            catch { return ""; }
        }

        private ExcelHints BuildExcelHints()
        {
            var hints = new ExcelHints();
            try { hints.ProcessId = Process.GetCurrentProcess().Id; } catch { }

            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app != null)
                {
                    try { hints.Hwnd = new IntPtr(app.Hwnd); } catch { }
                    try { hints.WorkbookPath = app.ActiveWorkbook?.FullName; } catch { }
                    hints.SessionMarker = _sessionId;
                }
            }
            catch { }

            return hints;
        }
    }
}
