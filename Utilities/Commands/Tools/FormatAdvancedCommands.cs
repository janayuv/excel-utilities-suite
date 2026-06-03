using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Copy the format of the active cell to every other cell in the selection.</summary>
    [ExcelCommand]
    public sealed class CopyCellFormattingCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.CopyCellFormatting",
            Label = "Copy Cell Formatting",
            Screentip = "Copy Cell Formatting",
            Supertip = "Apply the active cell's formatting (font, fill, number format, borders) to every other selected cell.",
            ImageId = "CopyCellFormatting",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 35,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range src = ctx.App.ActiveCell;
            if (src == null) { MessageBox.Show("Click the cell whose formatting to copy.", Definition.Label); return; }
            src.Copy();
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                if (cell.Address == src.Address) continue;
                cell.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            }
            ctx.App.CutCopyMode = (Excel.XlCutCopyMode)0; // cancel marquee
            MessageBox.Show("Formatting copied.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Apply alternating row shading (zebra stripes) to the selection.</summary>
    [ExcelCommand]
    public sealed class AlternateRowHighlightCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Format.AlternateRows",
            Label = "Alternate Row Colors",
            Screentip = "Alternate Row Highlight",
            Supertip = "Apply alternating fill colours to rows in the selection for easier reading (zebra-stripe effect).",
            ImageId = "AlternateRows",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 36,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new AlternateRowDialog();
    }

    internal sealed class AlternateRowDialog : Dialogs.DialogBase
    {
        private readonly Button _color1Btn = new Button { Text = "Color 1" };
        private readonly Button _color2Btn = new Button { Text = "Color 2" };
        private Color _color1 = Color.FromArgb(0xDD, 0xE5, 0xF0);
        private Color _color2 = Color.White;

        public AlternateRowDialog()
        {
            Text = AlternateRowHighlightCommand.Def.Label;
            ClientSize = new Size(300, 120);
            _color1Btn.BackColor = _color1;
            _color2Btn.BackColor = _color2;

            var lbl1 = new Label { Text = "Odd rows:", Left = 12, Top = 18, AutoSize = true };
            _color1Btn.SetBounds(90, 14, 80, 26);
            var lbl2 = new Label { Text = "Even rows:", Left = 12, Top = 52, AutoSize = true };
            _color2Btn.SetBounds(90, 48, 80, 26);

            _color1Btn.Click += (s, e) => { using (var d = new ColorDialog()) { if (d.ShowDialog() == DialogResult.OK) { _color1 = d.Color; _color1Btn.BackColor = _color1; } } };
            _color2Btn.Click += (s, e) => { using (var d = new ColorDialog()) { if (d.ShowDialog() == DialogResult.OK) { _color2 = d.Color; _color2Btn.BackColor = _color2; } } };

            var apply = new Button { Text = "&Apply", Left = 106, Top = 82, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 82, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                Color c1 = _color1, c2 = _color2;
                bool ok = RunOperation(AlternateRowHighlightCommand.Def, CurrentSelection, ctx =>
                {
                    int fRow = ctx.Target.Row, fCol = ctx.Target.Column;
                    int rows = ctx.Target.Rows.Count, cols = ctx.Target.Columns.Count;
                    Excel.Worksheet ws = ctx.Target.Worksheet;
                    for (int r = 0; r < rows; r++)
                    {
                        Excel.Range rowRange = ws.Range[ws.Cells[fRow + r, fCol], ws.Cells[fRow + r, fCol + cols - 1]];
                        rowRange.Interior.Color = ColorTranslator.ToOle(r % 2 == 0 ? c1 : c2);
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lbl1, _color1Btn, lbl2, _color2Btn, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Change the date display format of selected date cells.</summary>
    [ExcelCommand]
    public sealed class ConvertDateFormatCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Convert.DateFormat",
            Label = "Convert Date Format",
            Screentip = "Convert Date Format",
            Supertip = "Apply a different date display format to selected cells containing dates.",
            ImageId = "ConvertDateFormat",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 25,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new ConvertDateFormatDialog();
    }

    internal sealed class ConvertDateFormatDialog : Dialogs.DialogBase
    {
        private readonly ComboBox _fmt = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };

        public ConvertDateFormatDialog()
        {
            Text = ConvertDateFormatCommand.Def.Label;
            ClientSize = new Size(320, 100);
            var lbl = new Label { Text = "Format:", Left = 12, Top = 18, AutoSize = true };
            _fmt.Items.AddRange(new object[] {
                "dd/mm/yyyy","mm/dd/yyyy","yyyy-mm-dd","dd-mmm-yyyy",
                "d mmmm yyyy","mmmm d, yyyy","dd.mm.yyyy","yyyy/mm/dd"
            });
            _fmt.SelectedIndex = 0;
            _fmt.SetBounds(70, 15, 232, 23);
            var apply = new Button { Text = "&Apply", Left = 126, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 212, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string fmt = _fmt.Text;
                bool ok = RunOperation(ConvertDateFormatCommand.Def, CurrentSelection, ctx => { ctx.Target.NumberFormat = fmt; });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lbl, _fmt, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Swap the content of two same-sized ranges.</summary>
    [ExcelCommand]
    public sealed class SwapRangesCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Range.SwapRanges",
            Label = "Swap Ranges",
            Screentip = "Swap Two Ranges",
            Supertip = "Exchange the values of two same-sized ranges without a temporary helper column.",
            ImageId = "SwapRanges",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 70,
            RequiresSelection = false,
            UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SwapRangesDialog();
    }

    internal sealed class SwapRangesDialog : Dialogs.DialogBase
    {
        private readonly TextBox _rangeA = new TextBox { Text = "A1:A5" };
        private readonly TextBox _rangeB = new TextBox { Text = "C1:C5" };

        public SwapRangesDialog()
        {
            Text = SwapRangesCommand.Def.Label;
            ClientSize = new Size(320, 130);
            var lblA = new Label { Text = "Range A:", Left = 12, Top = 18, AutoSize = true };
            _rangeA.SetBounds(75, 15, 228, 23);
            var lblB = new Label { Text = "Range B:", Left = 12, Top = 50, AutoSize = true };
            _rangeB.SetBounds(75, 47, 228, 23);
            var apply = new Button { Text = "&Swap", Left = 126, Top = 90, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 212, Top = 90, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnSwap;
            Controls.AddRange(new Control[] { lblA, _rangeA, lblB, _rangeB, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _rangeA;
        }

        private void OnSwap(object sender, EventArgs e)
        {
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Excel.Range a, b;
            try { a = ws.Range[_rangeA.Text]; } catch { SetError(_rangeA, "Invalid range."); return; }
            try { b = ws.Range[_rangeB.Text]; } catch { SetError(_rangeB, "Invalid range."); return; }
            if (a.Rows.Count != b.Rows.Count || a.Columns.Count != b.Columns.Count)
            { MessageBox.Show("Ranges must be the same size.", SwapRangesCommand.Def.Label); return; }

            bool ok = RunOperation(SwapRangesCommand.Def, a, ctx =>
            {
                object[,] vA = a.Value2 as object[,];
                object[,] vB = b.Value2 as object[,];
                if (vA == null) { object sa = a.Value2; a.Value2 = b.Value2; b.Value2 = sa; return; }
                a.Value2 = vB;
                b.Value2 = vA;
            });
            if (ok) Close();
        }
    }

    /// <summary>Apply a currency number format to selected cells.</summary>
    [ExcelCommand]
    public sealed class ApplyCurrencyFormatCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Format.Currency",
            Label = "Currency Format",
            Screentip = "Apply Currency Format",
            Supertip = "Format selected cells as currency with a chosen symbol and decimal places.",
            ImageId = "CurrencyFormat",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 37,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new CurrencyFormatDialog();
    }

    internal sealed class CurrencyFormatDialog : Dialogs.DialogBase
    {
        private readonly ComboBox _symbol = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly NumericUpDown _dp = new NumericUpDown { Minimum = 0, Maximum = 6, Value = 2 };

        public CurrencyFormatDialog()
        {
            Text = ApplyCurrencyFormatCommand.Def.Label;
            ClientSize = new Size(290, 130);
            var lblS = new Label { Text = "Symbol:", Left = 12, Top = 18, AutoSize = true };
            _symbol.Items.AddRange(new object[] { "$ (USD)", "€ (EUR)", "£ (GBP)", "¥ (JPY)", "₹ (INR)", "₩ (KRW)" });
            _symbol.SelectedIndex = 0;
            _symbol.SetBounds(70, 15, 200, 23);
            var lblD = new Label { Text = "Decimals:", Left = 12, Top = 52, AutoSize = true };
            _dp.SetBounds(80, 49, 60, 23);
            var apply = new Button { Text = "&Apply", Left = 96, Top = 90, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 182, Top = 90, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string[] syms = { "$", "€", "£", "¥", "₹", "₩" };
                string sym = syms[_symbol.SelectedIndex];
                string dec = new string('0', (int)_dp.Value);
                string fmt = "\"" + sym + "\"#,##0" + (dec.Length > 0 ? "." + dec : "");
                bool ok = RunOperation(ApplyCurrencyFormatCommand.Def, CurrentSelection, ctx => { ctx.Target.NumberFormat = fmt; });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lblS, _symbol, lblD, _dp, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Convert between units (length, weight, temperature, area, volume) in-cell.</summary>
    [ExcelCommand]
    public sealed class UnitConversionCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Convert.Units",
            Label = "Unit Conversion",
            Screentip = "Unit Conversion",
            Supertip = "Convert cell values between common units: length, weight, temperature, area, and volume.",
            ImageId = "UnitConversion",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 40,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new UnitConversionDialog();
    }

    internal sealed class UnitConversionDialog : Dialogs.DialogBase
    {
        private static readonly string[] Units = {
            "km","m","cm","mm","mile","yard","foot","inch",
            "kg","g","mg","lb","oz",
            "C","F","K",
            "m2","km2","cm2","ha","acre",
            "L","mL","gallon","pint"
        };

        private readonly ComboBox _from = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox _to   = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };

        public UnitConversionDialog()
        {
            Text = UnitConversionCommand.Def.Label;
            ClientSize = new Size(320, 120);
            var lblF = new Label { Text = "From:", Left = 12, Top = 18, AutoSize = true };
            _from.Items.AddRange(Units); _from.SelectedIndex = 0;
            _from.SetBounds(55, 15, 100, 23);
            var lblT = new Label { Text = "To:", Left = 170, Top = 18, AutoSize = true };
            _to.Items.AddRange(Units); _to.SelectedIndex = 1;
            _to.SetBounds(192, 15, 100, 23);
            var apply = new Button { Text = "&Convert", Left = 106, Top = 76, Width = 90 };
            var cancel = new Button { Text = "&Cancel", Left = 202, Top = 76, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string from = _from.Text, to2 = _to.Text;
                bool ok = RunOperation(UnitConversionCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        object v = cell.Value2;
                        double d;
                        if (v == null || !double.TryParse(v.ToString(), out d)) continue;
                        double r = FromSI(ToSI(d, from), to2);
                        if (!double.IsNaN(r)) cell.Value2 = Math.Round(r, 6);
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lblF, _from, lblT, _to, apply, cancel });
            WireButtons(apply, cancel);
        }

        private static double ToSI(double v, string u)
        {
            switch (u)
            {
                case "km": return v * 1000; case "m": return v; case "cm": return v / 100; case "mm": return v / 1000;
                case "mile": return v * 1609.344; case "yard": return v * 0.9144; case "foot": return v * 0.3048; case "inch": return v * 0.0254;
                case "kg": return v; case "g": return v / 1000; case "mg": return v / 1e6; case "lb": return v * 0.453592; case "oz": return v * 0.0283495;
                case "C": return v + 273.15; case "F": return (v - 32) * 5.0 / 9 + 273.15; case "K": return v;
                case "m2": return v; case "km2": return v * 1e6; case "cm2": return v / 1e4; case "ha": return v * 1e4; case "acre": return v * 4046.86;
                case "L": return v / 1000; case "mL": return v / 1e6; case "gallon": return v * 0.00378541; case "pint": return v * 0.000473176;
                default: return double.NaN;
            }
        }

        private static double FromSI(double v, string u)
        {
            switch (u)
            {
                case "km": return v / 1000; case "m": return v; case "cm": return v * 100; case "mm": return v * 1000;
                case "mile": return v / 1609.344; case "yard": return v / 0.9144; case "foot": return v / 0.3048; case "inch": return v / 0.0254;
                case "kg": return v; case "g": return v * 1000; case "mg": return v * 1e6; case "lb": return v / 0.453592; case "oz": return v / 0.0283495;
                case "C": return v - 273.15; case "F": return (v - 273.15) * 9.0 / 5 + 32; case "K": return v;
                case "m2": return v; case "km2": return v / 1e6; case "cm2": return v * 1e4; case "ha": return v / 1e4; case "acre": return v / 4046.86;
                case "L": return v * 1000; case "mL": return v * 1e6; case "gallon": return v / 0.00378541; case "pint": return v / 0.000473176;
                default: return double.NaN;
            }
        }
    }
}
