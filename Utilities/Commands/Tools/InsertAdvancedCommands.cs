using System;
using System.Collections.Generic;
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
        // ── Controls ────────────────────────────────────────────────────────
        private readonly TextBox     _find      = new TextBox();
        private readonly TextBox     _replace   = new TextBox();
        private readonly RadioButton _scopeAll  = new RadioButton { Text = "All worksheets",    Checked = true };
        private readonly RadioButton _scopeCur  = new RadioButton { Text = "Current worksheet" };
        private RadioButton          _scopeRng;
        private readonly ComboBox    _search    = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ComboBox    _lookIn    = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly CheckBox    _matchCase = new CheckBox { Text = "Match Case" };
        private readonly CheckBox    _matchAll  = new CheckBox { Text = "Match entire cell contents" };
        private readonly CheckBox    _matchWord = new CheckBox { Text = "Match whole word or phrase only", Checked = true };
        private readonly Button      _btnFindAll  = new Button { Text = "&Find all" };
        private readonly Button      _btnClose    = new Button { Text = "&Close",  DialogResult = DialogResult.Cancel };
        private readonly Button      _btnReplace  = new Button { Text = "&Replace" };
        private readonly Button      _btnReplAll  = new Button { Text = "Replace &All" };
        private readonly Button      _btnFold     = new Button { Text = "Fold out >>" };
        private readonly Label       _lblResults  = new Label  { Text = "Results found: 0", AutoSize = true };
        private readonly ListView    _grid        = new ListView
        {
            View = View.Details, FullRowSelect = true, GridLines = true,
            HideSelection = false, BorderStyle = BorderStyle.FixedSingle
        };
        private readonly Panel       _resultsPanel = new Panel { Visible = false };

        private const int CollapsedH = 238;
        private const int ExpandedH  = 465;
        private bool _folded = true;
        private List<FindMatch> _matches = new List<FindMatch>();

        public FindReplaceAllSheetsDialog()
        {
            Text = FindReplaceAcrossSheetsCommand.Def.Label;
            ClientSize = new Size(660, CollapsedH);

            // Instruction text
            var lblHint = new Label
            {
                Text = "You can use ? (any single character) and * (any sequence).  " +
                       "Special codes:  {lf} = line feed  |  {cr} = carriage return  |  {tab} = tab",
                Left = 12, Top = 8, Width = 636, Height = 30, AutoSize = false
            };

            // Find / Replace
            var lblF = new Label { Text = "Find what:",    Left = 12, Top = 48, AutoSize = true };
            var lblR = new Label { Text = "Replace with:", Left = 12, Top = 76, AutoSize = true };
            _find.SetBounds(110, 45, 400, 23);
            _replace.SetBounds(110, 73, 400, 23);

            // Scope group box
            var gbScope = new GroupBox { Text = "Find/replace in:", Left = 12, Top = 106, Width = 206, Height = 100 };
            string selAddr = GetSelectionAddress();
            _scopeRng = new RadioButton { Text = "Range (" + selAddr + ")" };
            _scopeAll.SetBounds(6, 18, 195, 20);
            _scopeCur.SetBounds(6, 40, 195, 20);
            _scopeRng.SetBounds(6, 62, 195, 20);
            gbScope.Controls.AddRange(new Control[] { _scopeAll, _scopeCur, _scopeRng });

            // Options group box
            var gbOpts = new GroupBox { Text = "Options", Left = 228, Top = 106, Width = 424, Height = 100 };
            var lblSrc = new Label { Text = "Search:",  Left = 8, Top = 20, AutoSize = true };
            var lblLkn = new Label { Text = "Look in:", Left = 8, Top = 46, AutoSize = true };
            _search.Items.AddRange(new object[] { "By Rows", "By Columns" });
            _search.SelectedIndex = 0;
            _search.SetBounds(72, 17, 120, 23);
            _lookIn.Items.AddRange(new object[] { "Formulas", "Values" });
            _lookIn.SelectedIndex = 0;
            _lookIn.SetBounds(72, 43, 120, 23);
            _matchCase.SetBounds(8,  72, 200, 20);
            _matchWord.SetBounds(210, 17, 200, 20);
            _matchAll.SetBounds( 210, 39, 200, 20);
            gbOpts.Controls.AddRange(new Control[] { lblSrc, _search, lblLkn, _lookIn, _matchCase, _matchWord, _matchAll });

            // Button row
            int bTop = 176;
            _btnFindAll.SetBounds(12,  bTop, 88, 28);
            _btnClose.SetBounds(  106, bTop, 78, 28);
            _btnReplace.SetBounds(190, bTop, 80, 28);
            _btnReplAll.SetBounds(276, bTop, 90, 28);
            _btnFold.SetBounds(   560, bTop, 90, 28);

            _btnFindAll.Click += (s, e) => DoFindAll();
            _btnReplace.Click += (s, e) => DoReplaceSelected();
            _btnReplAll.Click += (s, e) => DoReplaceAll();
            _btnFold.Click    += (s, e) => ToggleFold();

            // Results panel (hidden until folded out)
            _resultsPanel.SetBounds(0, CollapsedH, 660, ExpandedH - CollapsedH);
            _lblResults.SetBounds(12, 8, 400, 20);
            _grid.SetBounds(12, 30, 636, ExpandedH - CollapsedH - 45);
            _grid.Columns.Add("Sheet",     100);
            _grid.Columns.Add("Cell",       60);
            _grid.Columns.Add("Value",     115);
            _grid.Columns.Add("Formula",   115);
            _grid.Columns.Add("Text",       95);
            _grid.Columns.Add("New Value", 110);
            _resultsPanel.Controls.AddRange(new Control[] { _lblResults, _grid });

            Controls.AddRange(new Control[] {
                lblHint, lblF, _find, lblR, _replace,
                gbScope, gbOpts,
                _btnFindAll, _btnClose, _btnReplace, _btnReplAll, _btnFold,
                _resultsPanel
            });
            WireButtons(_btnFindAll, _btnClose);
            ActiveControl = _find;
        }

        private void ToggleFold()
        {
            _folded = !_folded;
            _btnFold.Text = _folded ? "Fold out >>" : "<< Fold in";
            ClientSize = new Size(ClientSize.Width, _folded ? CollapsedH : ExpandedH);
            _resultsPanel.Visible = !_folded;
        }

        private void DoFindAll()
        {
            if (_find.TextLength == 0) { SetError(_find, "Enter text to find."); return; }
            SetError(_find, null);
            _matches = Search(_find.Text);
            PopulateGrid();
            if (_folded) ToggleFold();
        }

        private void DoReplaceSelected()
        {
            if (_find.TextLength == 0) { SetError(_find, "Enter text to find."); return; }
            SetError(_find, null);
            if (_matches.Count == 0) { DoFindAll(); return; }
            int idx = _grid.SelectedIndices.Count > 0 ? _grid.SelectedIndices[0] : 0;
            if (idx >= _matches.Count) return;
            ApplyReplacement(_matches[idx]);
            _matches.RemoveAt(idx);
            PopulateGrid();
        }

        private void DoReplaceAll()
        {
            if (_find.TextLength == 0) { SetError(_find, "Enter text to find."); return; }
            SetError(_find, null);
            var found = Search(_find.Text);
            int count = 0;
            foreach (var m in found) { ApplyReplacement(m); count++; }
            _matches.Clear();
            PopulateGrid();
            MessageBox.Show(count + " replacement(s) made.",
                FindReplaceAcrossSheetsCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private List<FindMatch> Search(string rawFind)
        {
            var results = new List<FindMatch>();
            bool mc      = _matchCase.Checked;
            bool mw      = _matchWord.Checked;
            bool me      = _matchAll.Checked;
            bool fmlas   = _lookIn.SelectedIndex == 0;
            bool byRows  = _search.SelectedIndex == 0;
            var cmp      = mc ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            string find  = Decode(rawFind);
            string rep   = Decode(_replace.Text);

            var xlAt     = me  ? Excel.XlLookAt.xlWhole : Excel.XlLookAt.xlPart;
            var xlLookIn = fmlas ? Excel.XlFindLookIn.xlFormulas : Excel.XlFindLookIn.xlValues;
            var xlOrder  = byRows ? Excel.XlSearchOrder.xlByRows : Excel.XlSearchOrder.xlByColumns;

            foreach (var ws in SheetsToSearch())
            {
                if (ws == null) continue;
                Excel.Range used;
                try { used = ws.UsedRange; } catch { continue; }

                Excel.Range first = null;
                try
                {
                    first = used.Find(find, Type.Missing, xlLookIn, xlAt, xlOrder,
                                      Excel.XlSearchDirection.xlNext, mc, Type.Missing, Type.Missing);
                }
                catch { }
                if (first == null) continue;

                string startAddr = first.Address;
                Excel.Range cur = first;
                do
                {
                    string val   = cur.Value2 != null ? cur.Value2.ToString() : "";
                    string fmla  = (cur.HasFormula as bool? == true)
                                   ? (cur.Formula as string ?? val) : val;
                    string text  = fmlas ? fmla : val;

                    // Post-filter for whole-word (Excel Find has no native whole-word option).
                    if (mw && !MatchesWholeWord(text, find, cmp)) goto next;

                    results.Add(new FindMatch
                    {
                        Sheet    = ws.Name,
                        Address  = cur.Address,
                        Value    = val,
                        Formula  = fmla,
                        NewValue = text.Replace(find, rep),
                        Cell     = cur
                    });
                    next:
                    try { cur = used.FindNext(cur); } catch { break; }
                } while (cur != null && cur.Address != startAddr);
            }
            return results;
        }

        private static bool MatchesWholeWord(string text, string find, StringComparison cmp)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(find)) return false;
            int idx = text.IndexOf(find, cmp);
            while (idx >= 0)
            {
                bool startOk = idx == 0 || !char.IsLetterOrDigit(text[idx - 1]);
                bool endOk   = idx + find.Length >= text.Length || !char.IsLetterOrDigit(text[idx + find.Length]);
                if (startOk && endOk) return true;
                idx = text.IndexOf(find, idx + 1, cmp);
            }
            return false;
        }

        private IEnumerable<Excel.Worksheet> SheetsToSearch()
        {
            var wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wb == null) yield break;
            if (_scopeAll.Checked)
            {
                foreach (Excel.Worksheet ws in wb.Worksheets) yield return ws;
            }
            else
            {
                yield return Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            }
        }

        private void ApplyReplacement(FindMatch m)
        {
            try
            {
                if (m.Cell == null) return;
                string find = Decode(_find.Text);
                string rep  = Decode(_replace.Text);
                bool fmlas  = _lookIn.SelectedIndex == 0;
                if (fmlas && (m.Cell.HasFormula as bool? == true))
                    m.Cell.Formula = (m.Cell.Formula as string ?? "").Replace(find, rep);
                else
                {
                    object v = m.Cell.Value2;
                    if (v != null) m.Cell.Value2 = v.ToString().Replace(find, rep);
                }
            }
            catch { }
        }

        private void PopulateGrid()
        {
            _grid.Items.Clear();
            foreach (var m in _matches)
            {
                var item = new ListViewItem(m.Sheet);
                item.SubItems.Add(m.Address);
                item.SubItems.Add(m.Value);
                item.SubItems.Add(m.Formula);
                item.SubItems.Add(m.Value);
                item.SubItems.Add(m.NewValue);
                _grid.Items.Add(item);
            }
            _lblResults.Text = "Results found: " + _matches.Count;
        }

        private static string GetSelectionAddress()
        {
            try
            {
                var rng = Globals.ThisAddIn.Application.Selection as Excel.Range;
                return rng != null ? rng.Address : "";
            }
            catch { return ""; }
        }

        private static string Decode(string s) => s
            .Replace("{lf}",  "\n")
            .Replace("{cr}",  "\r")
            .Replace("{tab}", "\t");

        private sealed class FindMatch
        {
            public string Sheet    { get; set; }
            public string Address  { get; set; }
            public string Value    { get; set; }
            public string Formula  { get; set; }
            public string NewValue { get; set; }
            public Excel.Range Cell { get; set; }
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
