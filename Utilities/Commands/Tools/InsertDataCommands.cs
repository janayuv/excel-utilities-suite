using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Fill blank cells downward with the value from the last non-blank cell above.</summary>
    [ExcelCommand]
    public sealed class FillDownBlanksCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.FillDownBlanks",
            Label = "Fill Down Blanks",
            Screentip = "Fill Blank Cells Down",
            Supertip = "Copy the nearest non-blank value above into every blank cell below it in the selection — ideal for un-merging filled reports.",
            ImageId = "FillDownBlanks",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 40,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int cols = ctx.Target.Columns.Count;
            int rows = ctx.Target.Rows.Count;
            int firstRow = ctx.Target.Row;
            int firstCol = ctx.Target.Column;
            Excel.Worksheet ws = ctx.Target.Worksheet;

            int filled = 0;
            for (int c = 1; c <= cols; c++)
            {
                object lastVal = null;
                for (int r = 1; r <= rows; r++)
                {
                    var cell = (Excel.Range)ws.Cells[firstRow + r - 1, firstCol + c - 1];
                    object v = cell.Value2;
                    if (v != null && v.ToString() != "")
                        lastVal = v;
                    else if (lastVal != null)
                    {
                        cell.Value2 = lastVal;
                        filled++;
                    }
                }
            }
            MessageBox.Show(filled + " blank cell(s) filled.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Transpose the selection in place (rows become columns, columns become rows).</summary>
    [ExcelCommand]
    public sealed class TransposeRangeCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.Transpose",
            Label = "Transpose Range",
            Screentip = "Transpose Range",
            Supertip = "Swap rows and columns in the selected range, pasting the transposed result starting at the top-left cell.",
            ImageId = "TransposeRange",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 50,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int rows = ctx.Target.Rows.Count;
            int cols = ctx.Target.Columns.Count;

            object[,] original = ctx.Target.Value2 as object[,];
            if (original == null)
            {
                MessageBox.Show("Select a range with at least 2 cells.", Definition.Label);
                return;
            }

            // Build transposed array (1-based).
            object[,] transposed = new object[cols + 1, rows + 1];
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    transposed[c, r] = original[r, c];

            Excel.Worksheet ws = ctx.Target.Worksheet;
            int startRow = ctx.Target.Row;
            int startCol = ctx.Target.Column;
            Excel.Range dest = ws.Range[
                ws.Cells[startRow, startCol],
                ws.Cells[startRow + cols - 1, startCol + rows - 1]];
            dest.Value2 = transposed;

            MessageBox.Show("Range transposed.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Select the cell containing the maximum value in the selection.</summary>
    [ExcelCommand]
    public sealed class SelectMaxCellCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.MaxCell",
            Label = "Select Max Cell",
            Screentip = "Select Maximum Value Cell",
            Supertip = "Jump to and select the cell containing the highest numeric value in the selection.",
            ImageId = "SelectMaxCell",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            try
            {
                double max = ctx.App.WorksheetFunction.Max(ctx.Target);
                foreach (Excel.Range cell in ctx.Target.Cells)
                {
                    object v = cell.Value2;
                    double d;
                    if (v != null && double.TryParse(v.ToString(), out d) && d == max)
                    {
                        cell.Select();
                        return;
                    }
                }
            }
            catch
            {
                MessageBox.Show("No numeric cells found.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>Select the cell containing the minimum value in the selection.</summary>
    [ExcelCommand]
    public sealed class SelectMinCellCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.MinCell",
            Label = "Select Min Cell",
            Screentip = "Select Minimum Value Cell",
            Supertip = "Jump to and select the cell containing the lowest numeric value in the selection.",
            ImageId = "SelectMinCell",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 31,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            try
            {
                double min = ctx.App.WorksheetFunction.Min(ctx.Target);
                foreach (Excel.Range cell in ctx.Target.Cells)
                {
                    object v = cell.Value2;
                    double d;
                    if (v != null && double.TryParse(v.ToString(), out d) && d == min)
                    {
                        cell.Select();
                        return;
                    }
                }
            }
            catch
            {
                MessageBox.Show("No numeric cells found.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>Concatenate all values in each column of the selection into a single delimited cell.</summary>
    [ExcelCommand]
    public sealed class CombineRowsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Formula.CombineRows",
            Label = "Combine Rows",
            Screentip = "Combine Row Values",
            Supertip = "Concatenate all cell values in each column of the selection into a single cell, separated by a delimiter you choose.",
            ImageId = "CombineRows",
            Tab = "Formula & Statistics",
            Group = "Combine",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new CombineRowsDialog();
    }

    internal sealed class CombineRowsDialog : Dialogs.DialogBase
    {
        private readonly TextBox _delim = new TextBox { Text = ", " };
        private readonly CheckBox _skipEmpty = new CheckBox { Text = "Skip empty cells", Checked = true };

        public CombineRowsDialog()
        {
            Text = CombineRowsCommand.Def.Label;
            ClientSize = new System.Drawing.Size(320, 130);

            var lblD = new Label { Text = "Delimiter:", Left = 12, Top = 18, AutoSize = true };
            _delim.SetBounds(80, 15, 220, 23);
            _skipEmpty.SetBounds(12, 46, 280, 22);

            var apply = new Button { Text = "&Combine", Left = 126, Top = 88, Width = 90 };
            var cancel = new Button { Text = "&Cancel", Left = 222, Top = 88, Width = 80, DialogResult = System.Windows.Forms.DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string delim = _delim.Text;
                bool skip = _skipEmpty.Checked;

                bool ok = RunOperation(CombineRowsCommand.Def, CurrentSelection, ctx =>
                {
                    Excel.Range area = ctx.Target;
                    int rows = area.Rows.Count;
                    int cols = area.Columns.Count;
                    int fRow = area.Row;
                    int fCol = area.Column;
                    Excel.Worksheet ws = area.Worksheet;

                    for (int c = 1; c <= cols; c++)
                    {
                        var parts = new List<string>();
                        for (int r = 1; r <= rows; r++)
                        {
                            object v = ((Excel.Range)ws.Cells[fRow + r - 1, fCol + c - 1]).Value2;
                            string str = v != null ? v.ToString() : "";
                            if (skip && str == "") continue;
                            parts.Add(str);
                        }
                        ((Excel.Range)ws.Cells[fRow, fCol + c - 1]).Value2 = string.Join(delim, parts);
                        for (int r = 2; r <= rows; r++)
                            ((Excel.Range)ws.Cells[fRow + r - 1, fCol + c - 1]).Value2 = null;
                    }
                });
                if (ok) Close();
            };

            Controls.AddRange(new System.Windows.Forms.Control[] { lblD, _delim, _skipEmpty, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _delim;
        }
    }
}
