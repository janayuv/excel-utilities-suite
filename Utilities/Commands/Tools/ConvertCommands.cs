using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>
    /// Convert text that looks like numbers into real numeric values. Port of the original
    /// DataConverterHelper.ConvertTextToNumbers, now guarded/undoable via the framework.
    /// </summary>
    [ExcelCommand]
    public sealed class ConvertTextToNumbersCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.TextToNumbers",
            Label = "Text to Numbers",
            Screentip = "Convert Text to Numbers",
            Supertip = "Turn numbers stored as text into real numeric values so they can be summed, sorted and used in formulas.",
            ImageId = "ConvertTextToNumbers",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int total = ctx.Target.Cells.Count;
            int done = 0;
            int converted = 0;

            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v != null)
                {
                    double number;
                    if (double.TryParse(v.ToString(), out number))
                    {
                        cell.Value2 = number;
                        converted++;
                    }
                }
                done++;
                if ((done & 0x3FF) == 0) ctx.Progress.Report((double)done / total);
            }

            ctx.Progress.Report(1);
            MessageBox.Show(converted + " value(s) converted to numbers.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>
    /// Convert numeric values into text (formatted as text, leading apostrophe). Port of the
    /// original DataConverterHelper.ConvertNumbersToText.
    /// </summary>
    [ExcelCommand]
    public sealed class ConvertNumbersToTextCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.NumbersToText",
            Label = "Numbers to Text",
            Screentip = "Convert Numbers to Text",
            Supertip = "Store numeric values as text (with a Text cell format) to preserve leading zeros and long IDs.",
            ImageId = "ConvertNumbersToText",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 11,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int total = ctx.Target.Cells.Count;
            int done = 0;
            int converted = 0;

            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v != null && v is double)
                {
                    cell.NumberFormat = "@";
                    cell.Value2 = "'" + v.ToString();
                    converted++;
                }
                done++;
                if ((done & 0x3FF) == 0) ctx.Progress.Report((double)done / total);
            }

            ctx.Progress.Report(1);
            MessageBox.Show(converted + " value(s) converted to text.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // ── Convert Time ─────────────────────────────────────────────────────────
    // Excel stores time as a fraction of a day: 1 hour = 1/24, 1 min = 1/1440, 1 sec = 1/86400.

    /// <summary>Convert time values to decimal hours (multiply by 24).</summary>
    [ExcelCommand]
    public sealed class ConvertTimeToHoursCommand : CommandBase
    {
        private const string _menuParent = "Convert Time";

        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.TimeToHours",
            Label = "Time to Hours",
            Screentip = "Time to Hours",
            Supertip = "Convert time values in the selection to decimal hours (e.g. 1:30 → 1.5).",
            ImageId = "ConvertDateFormat",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 20,
            MenuParent = _menuParent,
            MenuParentImageMso = "TimeGroupingHours",
            MenuParentOrder = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            ConvertTimeCells(ctx, 24, "hours");
        }

        internal static void ConvertTimeCells(CommandContext ctx, double factor, string unit)
        {
            int converted = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v is double d)
                {
                    cell.NumberFormat = "0.####";
                    cell.Value2 = d * factor;
                    converted++;
                }
            }
            MessageBox.Show(converted + " value(s) converted to " + unit + ".",
                "Convert Time", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Convert time values to decimal minutes (multiply by 1440).</summary>
    [ExcelCommand]
    public sealed class ConvertTimeToMinutesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.TimeToMinutes",
            Label = "Time to Minutes",
            Screentip = "Time to Minutes",
            Supertip = "Convert time values in the selection to decimal minutes (e.g. 1:30 → 90).",
            ImageId = "ConvertDateFormat",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 21,
            MenuParent = "Convert Time",
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx) =>
            ConvertTimeToHoursCommand.ConvertTimeCells(ctx, 1440, "minutes");
    }

    /// <summary>Convert time values to decimal seconds (multiply by 86400).</summary>
    [ExcelCommand]
    public sealed class ConvertTimeToSecondsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.TimeToSeconds",
            Label = "Time to Seconds",
            Screentip = "Time to Seconds",
            Supertip = "Convert time values in the selection to decimal seconds (e.g. 0:01:30 → 90).",
            ImageId = "ConvertDateFormat",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 22,
            MenuParent = "Convert Time",
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx) =>
            ConvertTimeToHoursCommand.ConvertTimeCells(ctx, 86400, "seconds");
    }

    /// <summary>Convert time values with a dialog (choose range, type, and optional destination).</summary>
    [ExcelCommand]
    public sealed class ConvertTimeCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Convert.TimeDialog",
            Label = "Convert Time...",
            Screentip = "Convert Time",
            Supertip = "Choose a range and conversion type (hours, minutes, or seconds), with an option to write results to a different location.",
            ImageId = "ConvertDateFormat",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 23,
            MenuParent = "Convert Time",
            RequiresSelection = false,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new ConvertTimeDialog();
    }

    internal sealed class ConvertTimeDialog : Dialogs.DialogBase
    {
        private readonly TextBox   _range    = new TextBox();
        private readonly RadioButton _toHours = new RadioButton { Text = "Time to hours",   Checked = true };
        private readonly RadioButton _toMins  = new RadioButton { Text = "Time to minutes" };
        private readonly RadioButton _toSecs  = new RadioButton { Text = "Time to seconds" };
        private readonly CheckBox   _saveTo   = new CheckBox { Text = "Save to another location (Convert range is one area)" };
        private readonly TextBox   _dest     = new TextBox { Enabled = false };

        public ConvertTimeDialog()
        {
            Text = ConvertTimeCommand.Def.Label.TrimEnd('.');
            ClientSize = new Size(360, 210);

            var lblRange = new Label { Text = "Convert Range:", Left = 12, Top = 18, AutoSize = true };
            _range.SetBounds(120, 15, 190, 23);
            _range.Text = RangeAddress();

            var btnPickRange = new Button { Text = "...", Left = 316, Top = 14, Width = 28, Height = 24 };
            btnPickRange.Click += (s, e) => PickRange();

            var lblType = new Label { Text = "Convert Type", Left = 12, Top = 52, AutoSize = true };
            lblType.Font = new Font(lblType.Font, FontStyle.Bold);
            _toHours.SetBounds(22, 70, 200, 20);
            _toMins.SetBounds(22, 92, 200, 20);
            _toSecs.SetBounds(22, 114, 200, 20);

            _saveTo.SetBounds(12, 142, 330, 20);
            _dest.SetBounds(12, 168, 200, 23);
            _saveTo.CheckedChanged += (s, e) => _dest.Enabled = _saveTo.Checked;

            var btnOK     = new Button { Text = "OK",     Left = 186, Top = 170, Width = 75 };
            var btnCancel = new Button { Text = "Cancel", Left = 268, Top = 170, Width = 75, DialogResult = DialogResult.Cancel };
            btnOK.Click += OnApply;

            Controls.AddRange(new Control[] {
                lblRange, _range, btnPickRange, lblType,
                _toHours, _toMins, _toSecs,
                _saveTo, _dest, btnOK, btnCancel
            });
            WireButtons(btnOK, btnCancel);
        }

        private static string RangeAddress()
        {
            try
            {
                var rng = Globals.ThisAddIn.Application.Selection as Excel.Range;
                return rng != null ? rng.Address : "";
            }
            catch { return ""; }
        }

        private void PickRange()
        {
            // Minimise and let the user select a range the standard way.
            // After refocus, read the current selection.
            this.WindowState = FormWindowState.Minimized;
            MessageBox.Show("Select the range in Excel, then click OK.", "Select Range",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.WindowState = FormWindowState.Normal;
            _range.Text = RangeAddress();
        }

        private void OnApply(object sender, EventArgs e)
        {
            string addr = _range.Text.Trim();
            if (string.IsNullOrEmpty(addr)) { SetError(_range, "Enter a range address."); return; }
            SetError(_range, null);

            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) { Close(); return; }

            Excel.Range src;
            try   { src = ws.Range[addr]; }
            catch { SetError(_range, "Invalid range address."); return; }

            double factor = _toHours.Checked ? 24 : _toMins.Checked ? 1440 : 86400;
            string unit   = _toHours.Checked ? "hours" : _toMins.Checked ? "minutes" : "seconds";
            bool   useDest = _saveTo.Checked && !string.IsNullOrWhiteSpace(_dest.Text);

            Excel.Range destStart = null;
            if (useDest)
            {
                try   { destStart = ws.Range[_dest.Text.Trim()]; }
                catch { SetError(_dest, "Invalid destination address."); return; }
            }

            int converted = 0, col = 1, row = 1;
            foreach (Excel.Range cell in src.Cells)
            {
                object v = cell.Value2;
                if (v is double d)
                {
                    if (useDest)
                    {
                        var target = destStart.Offset[row - 1, col - 1] as Excel.Range;
                        if (target != null) { target.NumberFormat = "0.####"; target.Value2 = d * factor; }
                    }
                    else
                    {
                        cell.NumberFormat = "0.####";
                        cell.Value2 = d * factor;
                    }
                    converted++;
                }
                col++;
                if (col > src.Columns.Count) { col = 1; row++; }
            }

            Close();
            MessageBox.Show(converted + " value(s) converted to " + unit + ".",
                ConvertTimeCommand.Def.Screentip, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
