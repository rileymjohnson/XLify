using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XLify
{
    internal static class RoslynWorkerClient
    {
        private static readonly object Gate = new object();
        private static dynamic _executor;

        private static dynamic EnsureExecutor()
        {
            if (_executor != null) return _executor;
            lock (Gate)
            {
                if (_executor != null) return _executor;
                var type = Type.GetTypeFromProgID("XLify.CSharpExecutor");
                if (type == null) throw new InvalidOperationException("COM server XLify.CSharpExecutor is not registered.");
                _executor = Activator.CreateInstance(type);
                return _executor;
            }
        }

        public static Task<ExecuteResponse> ExecuteAsync(
            string code,
            string sessionId,
            int timeoutMs = 15000,
            CancellationToken cancel = default,
            ExecutionOptions options = null,
            ExcelHints hints = null)
        {
            return Task.Run(() =>
            {
                var exec = EnsureExecutor();
                var result = exec.ExecuteInSession(sessionId ?? "default", code ?? string.Empty, timeoutMs);
                return new ExecuteResponse
                {
                    Success = result.Success,
                    Error = result.Error,
                    Output = result.Output,
                    TimedOut = result.TimedOut
                };
            }, cancel);
        }

        public static Task<SessionResponse> CreateSessionAsync(string sessionId, CancellationToken cancel = default)
        {
            return Task.Run(() =>
            {
                var exec = EnsureExecutor();
                exec.CreateSession(sessionId ?? "default", Globals.ThisAddIn?.Application);
                return new SessionResponse { SessionId = sessionId, Success = true };
            }, cancel);
        }

        public static Task DestroySessionAsync(string sessionId, CancellationToken cancel = default)
        {
            return Task.Run(() =>
            {
                var exec = EnsureExecutor();
                exec.DestroySession(sessionId ?? "default");
            }, cancel);
        }
    }

    internal sealed class ExecutionOptions
    {
        public int TimeoutMs { get; set; }
    }

    internal sealed class ExcelHints
    {
        public int ProcessId { get; set; }
        public IntPtr Hwnd { get; set; }
        public string SessionMarker { get; set; }
        public string WorkbookPath { get; set; }
    }

    internal sealed class CompilationError
    {
        public string Severity { get; set; }
        public int Line { get; set; }
        public int Column { get; set; }
        public string Message { get; set; }
    }

    internal sealed class ExecuteResponse
    {
        public bool Success { get; set; }
        public bool TimedOut { get; set; }
        public string Error { get; set; }
        public string Output { get; set; }
        public List<CompilationError> CompilationErrors { get; set; }
    }

    internal sealed class SessionResponse
    {
        public string SessionId { get; set; }
        public bool Success { get; set; }
    }
}
