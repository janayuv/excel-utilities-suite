using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Select only the duplicate cells in the selection (cells whose value appears more than once).</summary>
    [ExcelCommand]
    public sealed class SelectDuplicateCellsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.DuplicateCells",
            Label = "Select Duplicates",
            Screentip = "Select Duplicate Cells",
            Supertip = "Select every cell in the selection whose value appears more than once, so you can delete or format them.",
            ImageId = "SelectDuplicates",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 40,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                string k = v.ToString();
                counts[k] = counts.ContainsKey(k) ? counts[k] + 1 : 1;
            }

            Excel.Range result = null;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                int c;
                if (counts.TryGetValue(v.ToString(), out c) && c > 1)
                    result = result == null ? cell : ctx.App.Union(result, cell);
            }

            if (result != null) result.Select();
            else MessageBox.Show("No duplicate values found.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Select only the unique cells in the selection (cells whose value appears exactly once).</summary>
    [ExcelCommand]
    public sealed class SelectUniqueCellsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.UniqueCells",
            Label = "Select Uniques",
            Screentip = "Select Unique Cells",
            Supertip = "Select every cell in the selection whose value appears exactly once.",
            ImageId = "SelectUniques",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 41,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            var counts = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                string k = v.ToString();
                counts[k] = counts.ContainsKey(k) ? counts[k] + 1 : 1;
            }

            Excel.Range result = null;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                int c;
                if (counts.TryGetValue(v.ToString(), out c) && c == 1)
                    result = result == null ? cell : ctx.App.Union(result, cell);
            }

            if (result != null) result.Select();
            else MessageBox.Show("No unique values found.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Select cells matching a condition (equals, greater than, contains, etc.).</summary>
    [ExcelCommand]
    public sealed class SelectCellsByValueCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Select.ByValue",
            Label = "Select by Value",
            Screentip = "Select Cells by Value",
            Supertip = "Select cells in the current selection that match a condition: equals, greater than, less than, or contains.",
            ImageId = "SelectByValue",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 50,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SelectByValueDialog();
    }

    internal sealed class SelectByValueDialog : Dialogs.DialogBase
    {
        private readonly ComboBox _cond = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly TextBox _value = new TextBox();

        public SelectByValueDialog()
        {
            Text = SelectCellsByValueCommand.Def.Label;
            ClientSize = new Size(320, 120);

            _cond.Items.AddRange(new object[] {
                "Equals", "Not Equals", "Greater Than", "Less Than",
                "Greater Than or Equal", "Less Than or Equal", "Contains", "Does Not Contain"
            });
            _cond.SelectedIndex = 0;
            _cond.SetBounds(12, 12, 160, 23);

            var lblV = new Label { Text = "Value:", Left = 12, Top = 46, AutoSize = true };
            _value.SetBounds(60, 43, 240, 23);

            var apply = new Button { Text = "&Select", Left = 126, Top = 78, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 212, Top = 78, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnSelect;

            Controls.AddRange(new Control[] { _cond, lblV, _value, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _value;
        }

        private void OnSelect(object sender, EventArgs e)
        {
            if (_value.TextLength == 0 && _cond.SelectedIndex < 6)
            { SetError(_value, "Enter a value."); return; }
            SetError(_value, null);

            string condition = _cond.Text;
            string raw = _value.Text;
            double numVal;
            bool isNum = double.TryParse(raw, out numVal);

            Excel.Range result = null;
            foreach (Excel.Range cell in CurrentSelection.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                string sv = v.ToString();
                double cv;
                bool cellIsNum = double.TryParse(sv, out cv);
                bool match = false;

                switch (condition)
                {
                    case "Equals": match = string.Equals(sv, raw, StringComparison.OrdinalIgnoreCase); break;
                    case "Not Equals": match = !string.Equals(sv, raw, StringComparison.OrdinalIgnoreCase); break;
                    case "Greater Than": match = cellIsNum && isNum && cv > numVal; break;
                    case "Less Than": match = cellIsNum && isNum && cv < numVal; break;
                    case "Greater Than or Equal": match = cellIsNum && isNum && cv >= numVal; break;
                    case "Less Than or Equal": match = cellIsNum && isNum && cv <= numVal; break;
                    case "Contains": match = sv.IndexOf(raw, StringComparison.OrdinalIgnoreCase) >= 0; break;
                    case "Does Not Contain": match = sv.IndexOf(raw, StringComparison.OrdinalIgnoreCase) < 0; break;
                }

                if (match)
                    result = result == null ? cell : Globals.ThisAddIn.Application.Union(result, cell);
            }

            Close();
            if (result != null) result.Select();
            else MessageBox.Show("No matching cells found.", SelectCellsByValueCommand.Def.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Select all cells in the used range whose font colour matches the active cell.</summary>
    [ExcelCommand]
    public sealed class SelectByFontColorCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.ByFontColor",
            Label = "Select by Font Color",
            Screentip = "Select Cells by Font Color",
            Supertip = "Select all cells on the sheet whose font colour matches the active cell's font colour.",
            ImageId = "SelectByFontColor",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 11,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range active = ctx.App.ActiveCell;
            if (active == null) { MessageBox.Show("Click a cell with the font colour to match.", Definition.Label); return; }

            object targetColor = active.Font.Color;
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            Excel.Range used = ws != null ? ws.UsedRange : ctx.Target;

            Excel.Range matches = null;
            int total = used.Cells.Count, done = 0;
            foreach (Excel.Range cell in used.Cells)
            {
                if (Equals(cell.Font.Color, targetColor))
                    matches = matches == null ? cell : ctx.App.Union(matches, cell);
                if ((++done & 0x3FF) == 0) ctx.Progress.Report((double)done / total);
            }

            if (matches != null) matches.Select();
            else MessageBox.Show("No other cells share that font colour.", Definition.Label);
        }
    }
}
