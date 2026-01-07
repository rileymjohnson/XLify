using Serilog.Core;
using Serilog.Events;

namespace XLify
{
    // Adds Subsystem="SK" to events originating from Semantic Kernel or our SK-prefixed logs
    internal sealed class SemanticKernelTagEnricher : ILogEventEnricher
    {
        public void Enrich(LogEvent logEvent, ILogEventPropertyFactory propertyFactory)
        {
            if (logEvent == null || propertyFactory == null) return;

            bool isSK = false;

            // Check SourceContext if present
            if (logEvent.Properties != null && logEvent.Properties.TryGetValue("SourceContext", out var sc))
            {
                var s = sc.ToString().Trim('"');
                if (!string.IsNullOrEmpty(s) && s.StartsWith("Microsoft.SemanticKernel"))
                {
                    isSK = true;
                }
            }

            // Also tag our custom SK-* logs by message prefix
            if (!isSK)
            {
                var text = logEvent.MessageTemplate?.Text;
                if (!string.IsNullOrEmpty(text) && text.StartsWith("SK-"))
                {
                    isSK = true;
                }
            }

            if (isSK)
            {
                if (!logEvent.Properties.ContainsKey("Subsystem"))
                {
                    var prop = propertyFactory.CreateProperty("Subsystem", "SK");
                    logEvent.AddPropertyIfAbsent(prop);
                }
            }
        }
    }
}

