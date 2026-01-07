using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using Serilog;

namespace XLify
{
    internal static class LoggingBridge
    {
        private static bool _enabled;

        public static void EnableSeqForApp(string seqUrl = "http://localhost:5341")
        {
            if (_enabled) return;
            _enabled = true;
            try
            {
                // If Log.Logger wasn't configured yet, set a reasonable default to Seq
                try
                {
                    if (Log.Logger == Serilog.Core.Logger.None)
                    {
                        Log.Logger = new LoggerConfiguration()
                            .MinimumLevel.Debug()
                            .Enrich.FromLogContext()
                            .Enrich.With(new SemanticKernelTagEnricher())
                            .Enrich.WithProperty("App", "XLify.AddIn")
                            .Enrich.WithProperty("Workspace", "XLify-AddIn")
                            .WriteTo.Seq(seqUrl)
                            .CreateLogger();
                    }
                }
                catch { }

                // Bridge Console.Out and Console.Error to Serilog
                try { Console.SetOut(new SerilogTextWriter((msg) => Log.Information("[stdout] {Message}", msg))); } catch { }
                try { Console.SetError(new SerilogTextWriter((msg) => Log.Error("[stderr] {Message}", msg))); } catch { }

                // Bridge System.Diagnostics.Trace/Debug to Serilog
                try { Trace.Listeners.Add(new SerilogTraceListener()); } catch { }
            }
            catch { }
        }
    }

    internal sealed class SerilogTextWriter : TextWriter
    {
        private readonly Action<string> _emit;
        private readonly StringBuilder _buffer = new StringBuilder();

        public SerilogTextWriter(Action<string> emit)
        {
            _emit = emit ?? (_ => { });
        }

        public override Encoding Encoding => Encoding.UTF8;

        public override void Write(char value)
        {
            if (value == '\n')
            {
                FlushBuffer();
                return;
            }
            if (value != '\r') _buffer.Append(value);
        }

        public override void Write(string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            var lines = value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0) FlushBuffer();
                _buffer.Append(lines[i]);
            }
        }

        public override void WriteLine(string value)
        {
            Write(value);
            FlushBuffer();
        }

        public override void Flush()
        {
            FlushBuffer();
        }

        private void FlushBuffer()
        {
            try
            {
                if (_buffer.Length == 0) return;
                var s = _buffer.ToString();
                _buffer.Clear();
                _emit(s);
            }
            catch { }
        }
    }

    internal sealed class SerilogTraceListener : TraceListener
    {
        public override void Write(string message)
        {
            try { Log.Information("[trace] {Message}", message); } catch { }
        }

        public override void WriteLine(string message)
        {
            try { Log.Information("[trace] {Message}", message); } catch { }
        }

        public override void TraceEvent(TraceEventCache eventCache, string source, TraceEventType eventType, int id, string message)
        {
            try
            {
                switch (eventType)
                {
                    case TraceEventType.Critical:
                    case TraceEventType.Error:
                        Log.Error("[trace:{Source}] {Message}", source, message);
                        break;
                    case TraceEventType.Warning:
                        Log.Warning("[trace:{Source}] {Message}", source, message);
                        break;
                    case TraceEventType.Information:
                        Log.Information("[trace:{Source}] {Message}", source, message);
                        break;
                    default:
                        Log.Debug("[trace:{Source}] {Message}", source, message);
                        break;
                }
            }
            catch { }
        }
    }
}
