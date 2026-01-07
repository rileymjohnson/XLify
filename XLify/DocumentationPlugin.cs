using System;
using System.ComponentModel;
using System.Text;
using Microsoft.SemanticKernel;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLify
{
    internal sealed class DocumentationPlugin
    {
        [KernelFunction("generate_workbook_overview")]
        [Description("Return a Markdown overview of the active workbook: sheets and counts.")]
        public string GenerateWorkbookOverview([Description("Also write to a new sheet named 'Documentation'")] bool writeToNewSheet = false)
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return "No Excel Application available.";
            Excel.Workbook wb = null; try { wb = app.ActiveWorkbook; } catch { }
            if (wb == null) return "No active workbook.";

            var sb = new StringBuilder();
            sb.AppendLine("# Workbook Overview");
            sb.AppendLine($"- Name: {Safe(() => wb.Name)}");
            sb.AppendLine($"- Full Path: {Safe(() => wb.FullName)}");
            sb.AppendLine($"- Sheets: {Safe(() => wb.Worksheets?.Count, 0)}\n");
            sb.AppendLine("## Sheets");
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                sb.AppendLine($"- {Safe(() => ws.Name)}");
            }
            var md = sb.ToString();
            if (writeToNewSheet) TryWriteToNewSheet(wb, "Documentation", md);
            return md;
        }

        [KernelFunction("summarize_selection")]
        [Description("Return a Markdown summary of the current selection: address and size.")]
        public string SummarizeSelection([Description("Also write to a new sheet named 'Selection Summary'")] bool writeToNewSheet = false)
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return "No Excel Application available.";
            Excel.Range sel = null; try { sel = app.Selection as Excel.Range; } catch { }
            if (sel == null) return "No selection (or selection is not a range).";
            var addr = Safe(() => sel.Address[true, true, Excel.XlReferenceStyle.xlA1]);
            int rows = Safe(() => sel.Rows.Count, 0);
            int cols = Safe(() => sel.Columns.Count, 0);
            var md = $"# Selection Summary\n- Address: `{addr}`\n- Size: {rows} x {cols}\n";
            if (writeToNewSheet) TryWriteToNewSheet(app.ActiveWorkbook, "Selection Summary", md);
            return md;
        }

        private static void TryWriteToNewSheet(Excel.Workbook wb, string sheetName, string content)
        {
            if (wb == null || string.IsNullOrWhiteSpace(content)) return;
            try
            {
                var ws = wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]) as Excel.Worksheet;
                try { if (!string.IsNullOrWhiteSpace(sheetName)) ws.Name = sheetName; } catch { }
                var tgt = ws.Range["A1"] as Excel.Range;
                try { tgt.Value2 = content; tgt.WrapText = true; ws.Columns[1].ColumnWidth = 100; } catch { }
            }
            catch { }
        }

        private static T Safe<T>(Func<T> f, T fallback = default) { try { return f(); } catch { return fallback; } }
    }
}