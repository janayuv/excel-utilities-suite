using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Count the number of words in every selected cell and show a total.</summary>
    [ExcelCommand]
    public sealed class CountWordsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.CountWords",
            Label = "Count Words",
            Screentip = "Count Words in Selection",
            Supertip = "Count the total number of words across all selected cells.",
            ImageId = "CountWords",
            Tab = "Editing",
            Group = "Text",
            Order = 60,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            int totalWords = 0, cellsWithText = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                string s = v.ToString().Trim();
                if (s.Length == 0) continue;
                totalWords += Regex.Split(s, @"\s+").Length;
                cellsWithText++;
            }
            MessageBox.Show("Words: " + totalWords + "\nCells with text: " + cellsWithText,
                Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Pad values in the selection with leading zeros to a fixed width.</summary>
    [ExcelCommand]
    public sealed class AddLeadingZerosCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Text.AddLeadingZeros",
            Label = "Add Leading Zeros",
            Screentip = "Pad with Leading Zeros",
            Supertip = "Pad numeric or text values with leading zeros to reach a fixed total length (e.g. 42 becomes 00042).",
            ImageId = "AddLeadingZeros",
            Tab = "Editing",
            Group = "Text",
            Order = 61,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new AddLeadingZerosDialog();
    }

    internal sealed class AddLeadingZerosDialog : Dialogs.DialogBase
    {
        private readonly NumericUpDown _width = new NumericUpDown { Minimum = 1, Maximum = 50, Value = 5 };

        public AddLeadingZerosDialog()
        {
            Text = AddLeadingZerosCommand.Def.Label;
            ClientSize = new System.Drawing.Size(260, 100);
            var lbl = new Label { Text = "Total length:", Left = 12, Top = 18, AutoSize = true };
            _width.SetBounds(110, 15, 60, 23);
            var apply = new Button { Text = "&Apply", Left = 70, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 158, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                int w = (int)_width.Value;
                bool ok = RunOperation(AddLeadingZerosCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        object v = cell.Value2;
                        if (v == null) continue;
                        cell.NumberFormat = "@";
                        cell.Value2 = v.ToString().PadLeft(w, '0');
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lbl, _width, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Remove the leading apostrophe that Excel uses to force text storage.</summary>
    [ExcelCommand]
    public sealed class RemoveLeadingApostrophesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.RemoveApostrophes",
            Label = "Remove Apostrophes",
            Screentip = "Remove Leading Apostrophes",
            Supertip = "Strip the leading apostrophe that Excel uses to force text storage, then re-evaluate numeric values.",
            ImageId = "RemoveApostrophes",
            Tab = "Editing",
            Group = "Text",
            Order = 62,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int fixed2 = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                if (cell.PrefixCharacter == "'")
                {
                    object v = cell.Value2;
                    cell.NumberFormat = "General";
                    cell.Value2 = v;
                    fixed2++;
                }
            }
            MessageBox.Show(fixed2 + " apostrophe(s) removed.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Capitalise the first letter of every word in the selection (Proper Case).</summary>
    [ExcelCommand]
    public sealed class ProperCaseCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.ProperCase",
            Label = "Proper Case",
            Screentip = "Proper Case",
            Supertip = "Capitalise the first letter of every word in each selected cell.",
            ImageId = "ProperCase",
            Tab = "Editing",
            Group = "Text",
            Order = 63,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            var ti = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                cell.Value2 = ti.ToTitleCase(v.ToString().ToLower());
            }
        }
    }

    /// <summary>Replace a delimiter with a line break (Alt+Enter) in every selected cell.</summary>
    [ExcelCommand]
    public sealed class AddLineBreakCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Text.AddLineBreak",
            Label = "Add Line Break",
            Screentip = "Add Line Break",
            Supertip = "Replace a delimiter character with a line break (Alt+Enter) in each selected cell.",
            ImageId = "AddLineBreak",
            Tab = "Editing",
            Group = "Text",
            Order = 64,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new AddLineBreakDialog();
    }

    internal sealed class AddLineBreakDialog : Dialogs.DialogBase
    {
        private readonly TextBox _delim = new TextBox { Text = "," };

        public AddLineBreakDialog()
        {
            Text = AddLineBreakCommand.Def.Label;
            ClientSize = new System.Drawing.Size(300, 100);
            var lbl = new Label { Text = "Replace:", Left = 12, Top = 18, AutoSize = true };
            _delim.SetBounds(80, 15, 200, 23);
            var apply = new Button { Text = "&Apply", Left = 106, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                if (_delim.TextLength == 0) { SetError(_delim, "Enter delimiter."); return; }
                SetError(_delim, null);
                string delim = _delim.Text;
                bool ok = RunOperation(AddLineBreakCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        object v = cell.Value2;
                        if (v == null) continue;
                        cell.Value2 = v.ToString().Replace(delim, "\n");
                        cell.WrapText = true;
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lbl, _delim, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _delim;
        }
    }

    /// <summary>Convert digits in cells to Unicode superscript or subscript characters.</summary>
    [ExcelCommand]
    public sealed class SuperSubscriptCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Text.SuperSubscript",
            Label = "Super/Subscript",
            Screentip = "Superscript / Subscript",
            Supertip = "Convert digits in selected cells to Unicode superscript (x²) or subscript (x₂) characters.",
            ImageId = "SuperSubscript",
            Tab = "Editing",
            Group = "Text",
            Order = 65,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SuperSubscriptDialog();
    }

    internal sealed class SuperSubscriptDialog : Dialogs.DialogBase
    {
        private readonly RadioButton _super = new RadioButton { Text = "Superscript (x²)", Checked = true };
        private readonly RadioButton _sub = new RadioButton { Text = "Subscript (x₂)" };

        private static readonly string _digits    = "0123456789";
        private static readonly string _superUni  = "⁰¹²³⁴⁵⁶⁷⁸⁹";
        private static readonly string _subUni    = "₀₁₂₃₄₅₆₇₈₉";

        public SuperSubscriptDialog()
        {
            Text = SuperSubscriptCommand.Def.Label;
            ClientSize = new System.Drawing.Size(280, 120);
            _super.SetBounds(16, 16, 240, 22);
            _sub.SetBounds(16, 42, 240, 22);
            var apply = new Button { Text = "&Apply", Left = 86, Top = 80, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 172, Top = 80, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                bool sup = _super.Checked;
                string map = sup ? _superUni : _subUni;
                bool ok = RunOperation(SuperSubscriptCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        object v = cell.Value2;
                        if (v == null) continue;
                        var sb = new StringBuilder();
                        foreach (char c in v.ToString())
                        {
                            int idx = _digits.IndexOf(c);
                            sb.Append(idx >= 0 ? map[idx] : c);
                        }
                        cell.Value2 = sb.ToString();
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { _super, _sub, apply, cancel });
            WireButtons(apply, cancel);
        }
    }
}
