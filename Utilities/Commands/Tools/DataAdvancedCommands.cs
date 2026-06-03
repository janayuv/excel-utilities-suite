using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Wrap every formula in the selection with IFERROR to suppress error values.</summary>
    [ExcelCommand]
    public sealed class WrapIferrorCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Formula.WrapIferror",
            Label = "Wrap with IFERROR",
            Screentip = "Wrap Formulas with IFERROR",
            Supertip = "Add IFERROR around every formula in the selection so errors display as blank instead of #N/A, #VALUE! etc.",
            ImageId = "WrapIferror",
            Tab = "Formula & Statistics",
            Group = "Formulas",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FormulaOnly
        };

        protected override void Run(CommandContext ctx)
        {
            int wrapped = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                string formula = cell.Formula as string;
                if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=")) continue;
                string inner = formula.Substring(1);
                if (inner.StartsWith("IFERROR(", StringComparison.OrdinalIgnoreCase)) continue;
                cell.Formula = "=IFERROR(" + inner + ",\"\")";
                wrapped++;
            }
            MessageBox.Show(wrapped + " formula(s) wrapped with IFERROR.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Switch formula references between absolute, relative, and mixed modes.</summary>
    [ExcelCommand]
    public sealed class ToggleReferenceStyleCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Formula.ToggleRefStyle",
            Label = "Toggle Ref Style",
            Screentip = "Toggle Reference Style",
            Supertip = "Switch formula references between absolute ($A$1), relative (A1), or mixed ($A1 / A$1) in selected cells.",
            ImageId = "ToggleRefStyle",
            Tab = "Formula & Statistics",
            Group = "Formulas",
            Order = 25,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FormulaOnly
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new ToggleRefStyleDialog();
    }

    internal sealed class ToggleRefStyleDialog : Dialogs.DialogBase
    {
        private readonly RadioButton _abs    = new RadioButton { Text = "Absolute  ($A$1)" };
        private readonly RadioButton _rel    = new RadioButton { Text = "Relative  (A1)", Checked = true };
        private readonly RadioButton _mixRow = new RadioButton { Text = "Mixed - fix row  (A$1)" };
        private readonly RadioButton _mixCol = new RadioButton { Text = "Mixed - fix col  ($A1)" };

        public ToggleRefStyleDialog()
        {
            Text = ToggleReferenceStyleCommand.Def.Label;
            ClientSize = new System.Drawing.Size(280, 170);
            _abs.SetBounds(16, 12, 250, 22);
            _rel.SetBounds(16, 36, 250, 22);
            _mixRow.SetBounds(16, 60, 250, 22);
            _mixCol.SetBounds(16, 84, 250, 22);
            var apply = new Button { Text = "&Apply", Left = 86, Top = 126, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 172, Top = 126, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                bool absMode = _abs.Checked, relMode = _rel.Checked, mrMode = _mixRow.Checked;
                bool ok = RunOperation(ToggleReferenceStyleCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        string f = cell.Formula as string;
                        if (string.IsNullOrEmpty(f) || !f.StartsWith("=")) continue;
                        string converted = System.Text.RegularExpressions.Regex.Replace(
                            f, @"\$?([A-Z]+)\$?(\d+)", m =>
                            {
                                string col = m.Groups[1].Value, row = m.Groups[2].Value;
                                if (absMode) return "$" + col + "$" + row;
                                if (relMode) return col + row;
                                if (mrMode)  return col + "$" + row;
                                return "$" + col + row;
                            });
                        cell.Formula = converted;
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new System.Windows.Forms.Control[] { _abs, _rel, _mixRow, _mixCol, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Split delimited values in each cell into separate rows below.</summary>
    [ExcelCommand]
    public sealed class SplitIntoRowsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Data.SplitIntoRows",
            Label = "Split into Rows",
            Screentip = "Split Cell Values into Rows",
            Supertip = "Split each cell's delimited values into separate rows — e.g. 'A, B, C' becomes three rows.",
            ImageId = "SplitIntoRows",
            Tab = "Data & Cleaning",
            Group = "Split",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SplitIntoRowsDialog();
    }

    internal sealed class SplitIntoRowsDialog : Dialogs.DialogBase
    {
        private readonly TextBox _delim = new TextBox { Text = "," };

        public SplitIntoRowsDialog()
        {
            Text = SplitIntoRowsCommand.Def.Label;
            ClientSize = new System.Drawing.Size(300, 100);
            var lbl = new Label { Text = "Delimiter:", Left = 12, Top = 18, AutoSize = true };
            _delim.SetBounds(80, 15, 200, 23);
            var apply = new Button { Text = "&Split", Left = 106, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                if (_delim.TextLength == 0) { SetError(_delim, "Enter delimiter."); return; }
                SetError(_delim, null);
                string delim = _delim.Text;
                Excel.Range sel = CurrentSelection;
                var cells = new List<Excel.Range>();
                foreach (Excel.Range c in sel.Cells) cells.Add(c);
                cells.Sort((a2, b2) => b2.Row - a2.Row); // bottom-up so inserts don't shift indices

                bool ok = RunOperation(SplitIntoRowsCommand.Def, sel, ctx =>
                {
                    Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
                    foreach (Excel.Range cell in cells)
                    {
                        object v = cell.Value2;
                        if (v == null) continue;
                        string[] parts = v.ToString().Split(new string[] { delim }, StringSplitOptions.None);
                        if (parts.Length <= 1) continue;
                        cell.Value2 = parts[0].Trim();
                        for (int i = parts.Length - 1; i >= 1; i--)
                        {
                            ((Excel.Range)ws.Rows[cell.Row + 1]).Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
                            ((Excel.Range)ws.Cells[cell.Row + 1, cell.Column]).Value2 = parts[i].Trim();
                        }
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new System.Windows.Forms.Control[] { lbl, _delim, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _delim;
        }
    }

    /// <summary>Scan the selection and report cells with data-type mismatches.</summary>
    [ExcelCommand]
    public sealed class DetectDataTypesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Data.DetectTypes",
            Label = "Detect Data Types",
            Screentip = "Detect Data Type Issues",
            Supertip = "Scan the selection for type mismatches: numbers stored as text, dates stored as text, and hidden-character cells.",
            ImageId = "DetectDataTypes",
            Tab = "Data & Cleaning",
            Group = "Quality",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            int numAsText = 0, dateAsText = 0, hiddenChars = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                if (cell.PrefixCharacter == "'")
                {
                    object v = cell.Value2;
                    if (v == null) continue;
                    string s = v.ToString();
                    double d;
                    DateTime dt;
                    if (double.TryParse(s, out d)) numAsText++;
                    else if (DateTime.TryParse(s, out dt)) dateAsText++;
                }
                else if (cell.Value2 == null)
                {
                    object txt = cell.Text;
                    if (txt != null && txt.ToString().Trim().Length > 0) hiddenChars++;
                }
            }
            MessageBox.Show(
                "Data type scan:\n\nNumbers stored as text: " + numAsText +
                "\nDates stored as text: " + dateAsText +
                "\nHidden character cells: " + hiddenChars +
                (numAsText > 0 ? "\n\nTip: Use 'Text to Numbers' to fix numbers stored as text." : ""),
                Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Highlight near-duplicate rows using bigram similarity.</summary>
    [ExcelCommand]
    public sealed class FuzzyDedupeCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Data.FuzzyDedupe",
            Label = "Fuzzy Dedupe",
            Screentip = "Fuzzy Deduplication",
            Supertip = "Highlight rows in the selection that are similar (but not identical) to another row, based on a similarity threshold.",
            ImageId = "FuzzyDedupe",
            Tab = "Data & Cleaning",
            Group = "Duplicates",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new FuzzyDedupeDialog();
    }

    internal sealed class FuzzyDedupeDialog : Dialogs.DialogBase
    {
        private readonly NumericUpDown _threshold = new NumericUpDown { Minimum = 50, Maximum = 99, Value = 80 };

        public FuzzyDedupeDialog()
        {
            Text = FuzzyDedupeCommand.Def.Label;
            ClientSize = new System.Drawing.Size(300, 110);
            var lbl = new Label { Text = "Similarity %:", Left = 12, Top = 18, AutoSize = true };
            _threshold.SetBounds(110, 15, 70, 23);
            var hint = new Label { Text = "80% = flag rows that are 80% similar", Left = 12, Top = 48, AutoSize = true, ForeColor = System.Drawing.Color.Gray };
            var apply = new Button { Text = "&Scan", Left = 106, Top = 72, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 72, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                double thr = (double)_threshold.Value / 100.0;
                bool ok = RunOperation(FuzzyDedupeCommand.Def, CurrentSelection, ctx =>
                {
                    int rows = ctx.Target.Rows.Count, cols = ctx.Target.Columns.Count;
                    int fRow = ctx.Target.Row, fCol = ctx.Target.Column;
                    Excel.Worksheet ws = ctx.Target.Worksheet;
                    var rowStr = new string[rows];
                    for (int r = 0; r < rows; r++)
                    {
                        var sb = new StringBuilder();
                        for (int c = 0; c < cols; c++)
                        {
                            object v = ((Excel.Range)ws.Cells[fRow + r, fCol + c]).Value2;
                            sb.Append(v != null ? v.ToString() : ""); sb.Append('\t');
                        }
                        rowStr[r] = sb.ToString().ToLower();
                    }
                    int found = 0;
                    for (int i = 0; i < rows; i++)
                        for (int j = i + 1; j < rows; j++)
                        {
                            double sim = DiceSimilarity(rowStr[i], rowStr[j]);
                            if (sim >= thr && sim < 1.0)
                            {
                                ws.Range[ws.Cells[fRow + i, fCol], ws.Cells[fRow + i, fCol + cols - 1]].Interior.Color = Excel.XlRgbColor.rgbLightYellow;
                                ws.Range[ws.Cells[fRow + j, fCol], ws.Cells[fRow + j, fCol + cols - 1]].Interior.Color = Excel.XlRgbColor.rgbLightYellow;
                                found++;
                            }
                        }
                    MessageBox.Show(found > 0 ? found + " similar row pair(s) highlighted." : "No similar rows found.",
                        FuzzyDedupeCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
                });
                if (ok) Close();
            };
            Controls.AddRange(new System.Windows.Forms.Control[] { lbl, _threshold, hint, apply, cancel });
            WireButtons(apply, cancel);
        }

        private static double DiceSimilarity(string a, string b)
        {
            if (a == b) return 1.0;
            if (a.Length < 2 || b.Length < 2) return 0.0;
            var ba = new List<string>(); for (int i = 0; i < a.Length - 1; i++) ba.Add(a.Substring(i, 2));
            var bb = new List<string>(); for (int i = 0; i < b.Length - 1; i++) bb.Add(b.Substring(i, 2));
            return 2.0 * ba.Intersect(bb).Count() / (ba.Count + bb.Count);
        }
    }

    /// <summary>Calculate age in years from a birth-date column, writing results into the adjacent column.</summary>
    [ExcelCommand]
    public sealed class CalculateAgeCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Formula.CalculateAge",
            Label = "Calculate Age",
            Screentip = "Calculate Age from Date",
            Supertip = "Insert the age in years (as of today) into the column to the right of each selected date cell.",
            ImageId = "CalculateAge",
            Tab = "Formula & Statistics",
            Group = "Statistics",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            int inserted = 0;
            DateTime today = DateTime.Today;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                DateTime date;
                try { date = DateTime.FromOADate(System.Convert.ToDouble(v)); }
                catch { continue; }
                if (date > today) continue;
                int age = today.Year - date.Year;
                if (today < date.AddYears(age)) age--;
                cell.Offset[0, 1].Value2 = age;
                inserted++;
            }
            MessageBox.Show(inserted + " age(s) inserted.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
