using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Export the active sheet to a CSV file.</summary>
    [ExcelCommand]
    public sealed class ExportToCsvCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Export.ToCsv",
            Label = "Export to CSV",
            Screentip = "Export Sheet to CSV",
            Supertip = "Save the active worksheet as a comma-separated values (.csv) file.",
            ImageId = "ExportToCsv",
            Tab = "Export / Import",
            Group = "Export",
            Order = 30,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;

            using (var dlg = new SaveFileDialog
            {
                Title = "Export to CSV",
                Filter = "CSV file (*.csv)|*.csv",
                FileName = SanitizeName(ws.Name) + ".csv"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                Excel.Range used = ws.UsedRange;
                int rows = used.Rows.Count;
                int cols = used.Columns.Count;
                object[,] vals = used.Value2 as object[,];
                if (vals == null) return;

                var sb = new StringBuilder();
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        if (c > 1) sb.Append(',');
                        object v = vals[r, c];
                        string cell = v != null ? v.ToString() : "";
                        if (cell.Contains(",") || cell.Contains("\n") || cell.Contains("\""))
                            cell = "\"" + cell.Replace("\"", "\"\"") + "\"";
                        sb.Append(cell);
                    }
                    sb.AppendLine();
                }

                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show("Exported to:\n" + dlg.FileName, Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static string SanitizeName(string n)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) n = n.Replace(c, '_');
            return n;
        }
    }

    /// <summary>Export the active sheet to a JSON array file, using row 1 as property names.</summary>
    [ExcelCommand]
    public sealed class ExportToJsonCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Export.ToJson",
            Label = "Export to JSON",
            Screentip = "Export Sheet to JSON",
            Supertip = "Save the active worksheet as a JSON array, using the first row as property names.",
            ImageId = "ExportToJson",
            Tab = "Export / Import",
            Group = "Export",
            Order = 31,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;

            using (var dlg = new SaveFileDialog
            {
                Title = "Export to JSON",
                Filter = "JSON file (*.json)|*.json",
                FileName = SanitizeName(ws.Name) + ".json"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                Excel.Range used = ws.UsedRange;
                int rows = used.Rows.Count;
                int cols = used.Columns.Count;
                object[,] vals = used.Value2 as object[,];
                if (vals == null) return;

                string[] headers = new string[cols + 1];
                for (int c = 1; c <= cols; c++)
                    headers[c] = vals[1, c] != null ? vals[1, c].ToString() : "col" + c;

                var sb = new StringBuilder("[\n");
                for (int r = 2; r <= rows; r++)
                {
                    sb.Append("  {");
                    for (int c = 1; c <= cols; c++)
                    {
                        if (c > 1) sb.Append(", ");
                        object v = vals[r, c];
                        string val = v != null ? v.ToString() : "";
                        double d;
                        if (v != null && double.TryParse(val, out d))
                            sb.Append("\"" + JsonEscape(headers[c]) + "\": " + d.ToString("R"));
                        else
                            sb.Append("\"" + JsonEscape(headers[c]) + "\": \"" + JsonEscape(val) + "\"");
                    }
                    sb.Append(r < rows ? "}," : "}");
                    sb.AppendLine();
                }
                sb.Append("]");

                File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show("Exported to:\n" + dlg.FileName, Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static string JsonEscape(string s) =>
            s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\n", "\\n").Replace("\r", "\\r");

        private static string SanitizeName(string n)
        {
            foreach (char c in Path.GetInvalidFileNameChars()) n = n.Replace(c, '_');
            return n;
        }
    }

    /// <summary>Unmerge all merged cells in the selection and fill each resulting cell with the original value.</summary>
    [ExcelCommand]
    public sealed class UnmergeAndFillCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Data.UnmergeAndFill",
            Label = "Unmerge and Fill",
            Screentip = "Unmerge Cells and Fill",
            Supertip = "Split all merged cells in the selection and copy the original value into every resulting cell.",
            ImageId = "UnmergeAndFill",
            Tab = "Data & Cleaning",
            Group = "Merge",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int unmerged = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                if (!cell.MergeCells) continue;
                Excel.Range area = cell.MergeArea;
                object val = area.Cells[1, 1].Value2;
                area.UnMerge();
                area.Value2 = val;
                unmerged++;
            }
            MessageBox.Show(unmerged > 0 ? unmerged + " merged region(s) unmerged and filled."
                : "No merged cells found in the selection.",
                Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Highlight cells that differ between two same-sized ranges.</summary>
    [ExcelCommand]
    public sealed class CompareRangesCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Range.CompareRanges",
            Label = "Compare Ranges",
            Screentip = "Compare Two Ranges",
            Supertip = "Highlight cells that differ between two same-sized ranges — useful for comparing versions of a table.",
            ImageId = "CompareRanges",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 60,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new CompareRangesDialog();
    }

    internal sealed class CompareRangesDialog : Dialogs.DialogBase
    {
        private readonly TextBox _rangeA = new TextBox { Text = "A1:A10" };
        private readonly TextBox _rangeB = new TextBox { Text = "C1:C10" };

        public CompareRangesDialog()
        {
            Text = CompareRangesCommand.Def.Label;
            ClientSize = new System.Drawing.Size(340, 140);

            var lblA = new Label { Text = "Range A:", Left = 12, Top = 18, AutoSize = true };
            _rangeA.SetBounds(80, 15, 240, 23);
            var lblB = new Label { Text = "Range B:", Left = 12, Top = 50, AutoSize = true };
            _rangeB.SetBounds(80, 47, 240, 23);

            var apply = new Button { Text = "&Compare", Left = 136, Top = 98, Width = 90 };
            var cancel = new Button { Text = "&Cancel", Left = 232, Top = 98, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnCompare;

            Controls.AddRange(new System.Windows.Forms.Control[] { lblA, _rangeA, lblB, _rangeB, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _rangeA;
        }

        private void OnCompare(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet ws = app.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;

            Excel.Range a, b;
            try { a = ws.Range[_rangeA.Text]; } catch { SetError(_rangeA, "Invalid range."); return; }
            try { b = ws.Range[_rangeB.Text]; } catch { SetError(_rangeB, "Invalid range."); return; }

            if (a.Rows.Count != b.Rows.Count || a.Columns.Count != b.Columns.Count)
            {
                MessageBox.Show("Both ranges must be the same size.", CompareRangesCommand.Def.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int rows = a.Rows.Count, cols = a.Columns.Count, diffs = 0;
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                {
                    string sa = a.Cells[r, c].Value2?.ToString() ?? "";
                    string sbv = b.Cells[r, c].Value2?.ToString() ?? "";
                    if (sa != sbv)
                    {
                        a.Cells[r, c].Interior.Color = Excel.XlRgbColor.rgbLightCoral;
                        b.Cells[r, c].Interior.Color = Excel.XlRgbColor.rgbLightCoral;
                        diffs++;
                    }
                }

            Close();
            MessageBox.Show(diffs == 0 ? "Ranges are identical."
                : diffs + " difference(s) highlighted in coral.",
                CompareRangesCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
