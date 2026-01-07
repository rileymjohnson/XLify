using System;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using Microsoft.SemanticKernel.Agents.OpenAI;
using OpenAI;
using OpenAI.Responses;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Text.Json;

public class SeqTimerFilter : IFunctionInvocationFilter
{
    private static string Truncate(object value, int max = 1000)
    {
        if (value == null) return "<null>";
        var s = value.ToString() ?? "<null>";
        return s.Length <= max ? s : s.Substring(0, max) + " …(truncated)";
    }
    private static string Sha1Hex(string input)
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

    public async Task OnFunctionInvocationAsync(FunctionInvocationContext context, Func<FunctionInvocationContext, Task> next)
    {
        var plugin = context.Function?.Metadata?.PluginName ?? "<no-plugin>";
        var function = context.Function?.Name ?? "<no-function>";

        // Collect arguments (best-effort; KernelArguments implements IEnumerable)
        string argsStr = string.Empty;
        try
        {
            var parts = new System.Collections.Generic.List<string>();
            if (context.Arguments != null)
            {
                foreach (var kvp in context.Arguments)
                {
                    var key = string.IsNullOrWhiteSpace(kvp.Key) ? "<unnamed>" : kvp.Key;
                    if (string.Equals(key, "code", StringComparison.OrdinalIgnoreCase) && kvp.Value != null)
                    {
                        var codeStr = kvp.Value.ToString() ?? string.Empty;
                        var sha = Sha1Hex(codeStr);
                        var len = codeStr.Length;
                        var logCode = Environment.GetEnvironmentVariable("XLIFY_LOG_CODE");
                        if (!string.IsNullOrEmpty(logCode) && (logCode.Equals("1") || logCode.Equals("true", StringComparison.OrdinalIgnoreCase)))
                        {
                            parts.Add($"code=[{len} chars] sha1={sha} content={Truncate(codeStr, 500)}");
                        }
                        else
                        {
                            parts.Add($"code=[{len} chars] sha1={sha} (hidden)\u200b");
                        }
                    }
                    else
                    {
                        parts.Add($"{key}={Truncate(kvp.Value)}");
                    }
                }
            }
            argsStr = string.Join(", ", parts);
        }
        catch { /* ignore issues serializing args */ }

        Log.Information("SK-TOOL CALL: {Plugin}.{Function} args: {Args}", plugin, function, argsStr);

        var sw = Stopwatch.StartNew();
        object resultValue = null;
        bool threw = false;
        try
        {
            await next(context);
            try
            {
                // Capture a readable representation of the result
                if (context.Result != null)
                {
                    resultValue = context.Result.ToString();
                }
            }
            catch { /* ignore result extraction issues */ }
        }
        catch (Exception ex)
        {
            threw = true;
            sw.Stop();
            Log.Warning(ex, "SK-TOOL ERROR: {Plugin}.{Function} threw after {ElapsedMs} ms", plugin, function, sw.ElapsedMilliseconds);
            throw;
        }
        finally
        {
            sw.Stop();
        }

        if (!threw)
        {
            Log.Information(
                "SK-TOOL DONE: {Plugin}.{Function} in {ElapsedMs} ms. Result: {Result}",
                plugin, function, sw.ElapsedMilliseconds, Truncate(resultValue));
        }
    }
}

namespace XLify
{
    internal static class SemanticKernelFactory
    {
        public static Kernel CreateKernel(string sessionId)
        {
            var apiKey = ApiKeyVault.Get();
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                throw new InvalidOperationException("No API key found. Please save a key via ApiKeyVault (entered via WebView).");
            }

            var model = "gpt-5-mini";

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.FromLogContext()
                .Enrich.With(new SemanticKernelTagEnricher())
                .Enrich.WithProperty("App", "XLify.AddIn")
                .Enrich.WithProperty("Workspace", "XLify-AddIn")
                .WriteTo.Seq("http://localhost:5341") // Ensure Seq is running here
                .CreateLogger();

            // 2. Create the Logger Factory
            var loggerFactory = LoggerFactory.Create(builder =>
            {
                builder.AddSerilog(Log.Logger);
            });

            var builder = Kernel.CreateBuilder();

            // 3. Attach the LoggerFactory to the Kernel
            builder.Services.AddSingleton(loggerFactory);

            builder.AddOpenAIChatCompletion(model, apiKey, serviceId: "chat");

            // 4. Add the Performance Filter
            builder.Plugins.AddFromObject(new RoslynTool(sessionId), "roslyn");
            builder.Plugins.AddFromObject(new DocumentationPlugin(), "doc");

            // Register filters via DI so they’re discovered by SK
            try { builder.Services.AddSingleton<IPromptRenderFilter, SeqPromptFilter>(); } catch { }

            var kernel = builder.Build();

            // 5. Register the Timing Filter to catch bottlenecks
            kernel.FunctionInvocationFilters.Add(new SeqTimerFilter());
            // Ensure prompt filter is active even if DI discovery changes
            try { kernel.PromptRenderFilters.Add(new SeqPromptFilter()); } catch { }

            try
            {
                // Prepare an OpenAI Responses agent so downstream callers can opt-in to the Responses API.
                // The agent is created with the same API key and can be used with ChatHistoryChannel.
                var oaClient = new OpenAIClient(apiKey);
                // OPENAI001: Responses client factory is preview; suppress analyzer per SDK guidance
#pragma warning disable OPENAI001
                var responsesClient = oaClient.GetOpenAIResponseClient(model);
#pragma warning restore OPENAI001
                var responseAgent = new OpenAIResponseAgent(responsesClient);
                try { responseAgent.GetType().GetProperty("Kernel")?.SetValue(responseAgent, kernel); } catch { }

                // Stash the agent on kernel data for retrieval without changing method signatures.
                kernel.Data["__openai_response_agent__"] = responseAgent;
            }
            catch
            {
                // If the preview agents package is unavailable at runtime, keep kernel usable.
            }
            return kernel;
        }
    }
}

public class SeqPromptFilter : IPromptRenderFilter
{
    private static string Truncate(string s, int max = 2000)
    {
        if (string.IsNullOrEmpty(s)) return s ?? "";
        return s.Length <= max ? s : s.Substring(0, max) + " …(truncated)";
    }

    public async Task OnPromptRenderAsync(PromptRenderContext context, Func<PromptRenderContext, Task> next)
    {
        string template = null;
        try { template = context?.GetType().GetProperty("Template")?.GetValue(context)?.ToString(); } catch { }
        Log.Information("SK-PROMPT RENDER: template={Template}", Truncate(template ?? ""));

        await next(context);

        string rendered = null;
        try { rendered = context?.GetType().GetProperty("RenderedPrompt")?.GetValue(context)?.ToString(); } catch { }
        Log.Information("SK-PROMPT RENDERED: {Rendered}", Truncate(rendered ?? ""));
    }
}

// Note: AI service request/response filter omitted for SK 1.68,
// since the IAIServiceFilter hook is not available in this version.
