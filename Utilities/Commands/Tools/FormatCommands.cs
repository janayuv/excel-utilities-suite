using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Clear only the formatting (colours, fonts, borders) from the selection — keep values.</summary>
    [ExcelCommand]
    public sealed class ClearFormattingCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.ClearFormatting",
            Label = "Clear Formatting",
            Screentip = "Clear Formatting",
            Supertip = "Remove all cell formatting (fill colours, fonts, borders, number formats) from the selection without deleting the values.",
            ImageId = "ClearFormatting",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            ctx.Target.ClearFormats();
            MessageBox.Show("Formatting cleared.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Replace every formula in the selection with its current calculated value.</summary>
    [ExcelCommand]
    public sealed class ReplaceFormulasWithValuesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Formula.ReplaceWithValues",
            Label = "Formulas to Values",
            Screentip = "Replace Formulas with Values",
            Supertip = "Paste-as-values over every formula in the selection, locking in the current result and removing formula dependencies.",
            ImageId = "FormulasToValues",
            Tab = "Formula & Statistics",
            Group = "Formulas",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FormulaOnly
        };

        protected override void Run(CommandContext ctx)
        {
            object[,] vals = ctx.Target.Value2 as object[,];
            if (vals != null)
                ctx.Target.Value2 = vals;
            else
                ctx.Target.Value2 = ctx.Target.Value2;

            MessageBox.Show("Formulas replaced with values.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Round numeric values in the selection to a specified number of decimal places — in-cell, no formula.</summary>
    [ExcelCommand]
    public sealed class RoundValuesCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Format.RoundValues",
            Label = "Round Values",
            Screentip = "Round Values",
            Supertip = "Round numeric cell values to a chosen number of decimal places directly in-cell, without adding a ROUND formula.",
            ImageId = "RoundValues",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new RoundValuesDialog();
    }

    internal sealed class RoundValuesDialog : Dialogs.DialogBase
    {
        private readonly NumericUpDown _decimals = new NumericUpDown { Minimum = 0, Maximum = 15, Value = 2 };

        public RoundValuesDialog()
        {
            Text = RoundValuesCommand.Def.Label;
            ClientSize = new System.Drawing.Size(280, 110);

            var lbl = new Label { Text = "Decimal places:", Left = 12, Top = 18, AutoSize = true };
            _decimals.SetBounds(130, 15, 60, 23);

            var apply = new Button { Text = "&Apply", Left = 90, Top = 65, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 178, Top = 65, Width = 80, DialogResult = System.Windows.Forms.DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                int dp = (int)_decimals.Value;
                bool ok = RunOperation(RoundValuesCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        object v = cell.Value2;
                        double d;
                        if (v != null && double.TryParse(v.ToString(), out d))
                            cell.Value2 = Math.Round(d, dp, MidpointRounding.AwayFromZero);
                    }
                });
                if (ok) Close();
            };

            Controls.AddRange(new System.Windows.Forms.Control[] { lbl, _decimals, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _decimals;
        }
    }

    /// <summary>Change the sign of every numeric value in the selection (positive to negative and vice versa).</summary>
    [ExcelCommand]
    public sealed class ChangeSignCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.ChangeSign",
            Label = "Change Sign",
            Screentip = "Change Sign of Values",
            Supertip = "Flip positive numbers to negative and negative numbers to positive across the selection without helper formulas.",
            ImageId = "ChangeSign",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 21,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int changed = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                double d;
                if (v != null && double.TryParse(v.ToString(), out d))
                {
                    cell.Value2 = -d;
                    changed++;
                }
            }
            MessageBox.Show(changed + " value(s) sign-flipped.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Convert a number to its English word representation (e.g. 1234 to "One Thousand Two Hundred Thirty Four").</summary>
    [ExcelCommand]
    public sealed class NumbersToWordsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.NumbersToWords",
            Label = "Numbers to Words",
            Screentip = "Convert Numbers to Words",
            Supertip = "Replace numeric values with their English word equivalent — useful for cheque amounts and formal documents.",
            ImageId = "NumbersToWords",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 15,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                double d;
                if (v != null && double.TryParse(v.ToString(), out d))
                    cell.Value2 = NumberToWords((long)Math.Truncate(d));
            }
        }

        private static readonly string[] Ones = {
            "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten",
            "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen",
            "Eighteen", "Nineteen"
        };
        private static readonly string[] Tens = {
            "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"
        };

        private static string NumberToWords(long n)
        {
            if (n == 0) return "Zero";
            if (n < 0) return "Negative " + NumberToWords(-n);
            if (n < 20) return Ones[n];
            if (n < 100) return Tens[n / 10] + (n % 10 != 0 ? " " + Ones[n % 10] : "");
            if (n < 1000) return Ones[n / 100] + " Hundred" + (n % 100 != 0 ? " " + NumberToWords(n % 100) : "");
            if (n < 1_000_000) return NumberToWords(n / 1000) + " Thousand" + (n % 1000 != 0 ? " " + NumberToWords(n % 1000) : "");
            if (n < 1_000_000_000) return NumberToWords(n / 1_000_000) + " Million" + (n % 1_000_000 != 0 ? " " + NumberToWords(n % 1_000_000) : "");
            return NumberToWords(n / 1_000_000_000) + " Billion" + (n % 1_000_000_000 != 0 ? " " + NumberToWords(n % 1_000_000_000) : "");
        }
    }
}
