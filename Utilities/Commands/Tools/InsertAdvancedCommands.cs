using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Fill the selection with random integers, decimals, dates, or placeholder text.</summary>
    [ExcelCommand]
    public sealed class InsertRandomDataCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Insert.RandomData",
            Label = "Insert Random Data",
            Screentip = "Insert Random Data",
            Supertip = "Fill the selected cells with random integers, decimal numbers, dates, or Lorem-ipsum placeholder text.",
            ImageId = "InsertRandomData",
            Tab = "Insert",
            Group = "Insert",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new InsertRandomDataDialog();
    }

    internal sealed class InsertRandomDataDialog : Dialogs.DialogBase
    {
        private readonly ComboBox _type = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly NumericUpDown _min = new NumericUpDown { Minimum = -1000000, Maximum = 1000000, Value = 1 };
        private readonly NumericUpDown _max = new NumericUpDown { Minimum = -1000000, Maximum = 1000000, Value = 100 };
        private readonly Label _lblMin = new Label { Text = "Min:", AutoSize = true };
        private readonly Label _lblMax = new Label { Text = "Max:", AutoSize = true };

        public InsertRandomDataDialog()
        {
            Text = InsertRandomDataCommand.Def.Label;
            ClientSize = new Size(300, 145);

            var lblT = new Label { Text = "Type:", Left = 12, Top = 15, AutoSize = true };
            _type.Items.AddRange(new object[] { "Integer", "Decimal", "Date", "Lorem Text" });
            _type.SelectedIndex = 0;
            _type.SetBounds(60, 12, 220, 23);
            _type.SelectedIndexChanged += (s, e) =>
            {
                bool numeric = _type.SelectedIndex < 2;
                _lblMin.Visible = _min.Visible = _lblMax.Visible = _max.Visible = numeric;
            };

            _lblMin.SetBounds(12, 48, 30, 20);
            _min.SetBounds(50, 45, 90, 23);
            _lblMax.SetBounds(160, 48, 30, 20);
            _max.SetBounds(198, 45, 90, 23);

            var apply = new Button { Text = "&Insert", Left = 106, Top = 105, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 105, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnApply;

            Controls.AddRange(new Control[] { lblT, _type, _lblMin, _min, _lblMax, _max, apply, cancel });
            WireButtons(apply, cancel);
        }

        private void OnApply(object sender, EventArgs e)
        {
            string type = _type.Text;
            int minI = (int)_min.Value, maxI = (int)_max.Value;
            if ((type == "Integer" || type == "Decimal") && minI >= maxI)
            { SetError(_min, "Min must be less than Max."); return; }
            SetError(_min, null);

            var rnd = new Random();
            bool ok = RunOperation(InsertRandomDataCommand.Def, CurrentSelection, ctx =>
            {
                foreach (Excel.Range cell in ctx.Target.Cells)
                {
                    switch (type)
                    {
                        case "Integer":
                            cell.Value2 = rnd.Next(minI, maxI + 1);
                            break;
                        case "Decimal":
                            cell.Value2 = Math.Round(minI + rnd.NextDouble() * (maxI - minI), 2);
                            break;
                        case "Date":
                            var start = new DateTime(2000, 1, 1);
                            int days = (new DateTime(2025, 12, 31) - start).Days;
                            cell.Value2 = start.AddDays(rnd.Next(days)).ToOADate();
                            cell.NumberFormat = "dd/mm/yyyy";
                            break;
                        case "Lorem Text":
                            string[] words = { "lorem", "ipsum", "dolor", "sit", "amet", "consectetur",
                                "adipiscing", "elit", "sed", "do", "eiusmod", "tempor", "incididunt" };
                            int wc = rnd.Next(3, 8);
                            var sb = new StringBuilder();
                            for (int i = 0; i < wc; i++) { if (i > 0) sb.Append(' '); sb.Append(words[rnd.Next(words.Length)]); }
                            cell.Value2 = sb.ToString();
                            break;
                    }
                }
            });
            if (ok) Close();
        }
    }

    /// <summary>Find and replace text across every sheet in the workbook.</summary>
    [ExcelCommand]
    public sealed class FindReplaceAcrossSheetsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Text.FindReplaceAcrossSheets",
            Label = "Find & Replace All Sheets",
            Screentip = "Find and Replace Across All Sheets",
            Supertip = "Search every sheet in the workbook and replace text — not just the active sheet.",
            ImageId = "FindReplaceSheets",
            Tab = "Editing",
            Group = "Text",
            Order = 50,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new FindReplaceAllSheetsDialog();
    }

    internal sealed class FindReplaceAllSheetsDialog : Dialogs.DialogBase
    {
        private readonly TextBox _find = new TextBox();
        private readonly TextBox _replace = new TextBox();
        private readonly CheckBox _matchCase = new CheckBox { Text = "Match case" };

        public FindReplaceAllSheetsDialog()
        {
            Text = FindReplaceAcrossSheetsCommand.Def.Label;
            ClientSize = new Size(360, 148);

            var lblF = new Label { Text = "Find:", Left = 12, Top = 18, AutoSize = true };
            _find.SetBounds(70, 15, 270, 23);
            var lblR = new Label { Text = "Replace:", Left = 12, Top = 50, AutoSize = true };
            _replace.SetBounds(70, 47, 270, 23);
            _matchCase.SetBounds(70, 78, 180, 22);

            var apply = new Button { Text = "&Replace All", Left = 160, Top = 108, Width = 100 };
            var cancel = new Button { Text = "&Cancel", Left = 266, Top = 108, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnReplace;

            Controls.AddRange(new Control[] { lblF, _find, lblR, _replace, _matchCase, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _find;
        }

        private void OnReplace(object sender, EventArgs e)
        {
            if (_find.TextLength == 0) { SetError(_find, "Enter text to find."); return; }
            SetError(_find, null);

            string find = _find.Text, rep = _replace.Text;
            bool mc = _matchCase.Checked;
            int total = 0;

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wb == null) { Close(); return; }

            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                Excel.Range found = ws.UsedRange.Find(find,
                    Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext,
                    mc, Type.Missing, Type.Missing);

                if (found == null) continue;
                string firstAddr = found.Address;
                do
                {
                    object v = found.Value2;
                    if (v != null) found.Value2 = v.ToString().Replace(find, rep);
                    total++;
                    found = ws.UsedRange.FindNext(found);
                } while (found != null && found.Address != firstAddr);
            }

            Close();
            MessageBox.Show(total + " replacement(s) made across all sheets.",
                FindReplaceAcrossSheetsCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Flip a range vertically (reverse row order).</summary>
    [ExcelCommand]
    public sealed class FlipVerticalCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.FlipVertical",
            Label = "Flip Vertical",
            Screentip = "Flip Range Vertically",
            Supertip = "Reverse the row order of the selected range — the last row becomes the first.",
            ImageId = "FlipVertical",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 55,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int rows = ctx.Target.Rows.Count, cols = ctx.Target.Columns.Count;
            object[,] vals = ctx.Target.Value2 as object[,];
            if (vals == null) return;

            object[,] flipped = new object[rows + 1, cols + 1];
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    flipped[rows - r + 1, c] = vals[r, c];

            ctx.Target.Value2 = flipped;
            MessageBox.Show("Range flipped vertically.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Flip a range horizontally (reverse column order).</summary>
    [ExcelCommand]
    public sealed class FlipHorizontalCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.FlipHorizontal",
            Label = "Flip Horizontal",
            Screentip = "Flip Range Horizontally",
            Supertip = "Reverse the column order of the selected range — the last column becomes the first.",
            ImageId = "FlipHorizontal",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 56,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int rows = ctx.Target.Rows.Count, cols = ctx.Target.Columns.Count;
            object[,] vals = ctx.Target.Value2 as object[,];
            if (vals == null) return;

            object[,] flipped = new object[rows + 1, cols + 1];
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    flipped[r, cols - c + 1] = vals[r, c];

            ctx.Target.Value2 = flipped;
            MessageBox.Show("Range flipped horizontally.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Prefix every selected cell with a bullet or sequential number.</summary>
    [ExcelCommand]
    public sealed class InsertBulletsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Insert.Bullets",
            Label = "Insert Bullets",
            Screentip = "Insert Bullet Points",
            Supertip = "Prefix every selected cell with a bullet character of your choice, or number them sequentially.",
            ImageId = "InsertBullets",
            Tab = "Insert",
            Group = "Insert",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new InsertBulletsDialog();
    }

    internal sealed class InsertBulletsDialog : Dialogs.DialogBase
    {
        private readonly ComboBox _bullet = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly CheckBox _numbered = new CheckBox { Text = "Number sequentially instead" };

        public InsertBulletsDialog()
        {
            Text = InsertBulletsCommand.Def.Label;
            ClientSize = new Size(300, 120);

            var lbl = new Label { Text = "Bullet:", Left = 12, Top = 15, AutoSize = true };
            _bullet.Items.AddRange(new object[] { "* (bullet)", "- (dash)", "* (asterisk)", "> (arrow)" });
            _bullet.SelectedIndex = 0;
            _bullet.SetBounds(65, 12, 220, 23);
            _numbered.SetBounds(12, 44, 270, 22);

            var apply = new Button { Text = "&Apply", Left = 106, Top = 80, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 80, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string[] chars = { "• ", "- ", "* ", "> " };
                string ch = chars[_bullet.SelectedIndex];
                bool num = _numbered.Checked;
                int idx = 1;

                bool ok = RunOperation(InsertBulletsCommand.Def, CurrentSelection, ctx =>
                {
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        string existing = cell.Value2 != null ? cell.Value2.ToString() : "";
                        cell.Value2 = num ? idx++ + ". " + existing : ch + existing;
                    }
                });
                if (ok) Close();
            };

            Controls.AddRange(new Control[] { lbl, _bullet, _numbered, apply, cancel });
            WireButtons(apply, cancel);
        }
    }
}
