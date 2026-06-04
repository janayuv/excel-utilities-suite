using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using utilities.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    [ExcelCommand]
    public sealed class RegexFindReplaceCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id          = "Text.RegexFindReplace",
            Label       = "Regex Find & Replace",
            Screentip   = "Regex Find & Replace",
            Supertip    = "Find and replace cell text using regular expression patterns. Supports capture groups ($1, $2) in the replacement string.",
            ImageMso    = "ReplaceDialog",
            Tab         = "Editing",
            Group       = "Text",
            Order       = 55,
            RequiresSelection = false,
            UndoMode    = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override DialogBase CreateDialog() => new RegexFindReplaceDialog();
    }

    // ─────────────────────────────────────────────────────────────────────────

    internal sealed class RegexFindReplaceDialog : DialogBase
    {
        private readonly TextBox     _find      = new TextBox();
        private readonly TextBox     _replace   = new TextBox();
        private readonly RadioButton _rbSel     = new RadioButton { Text = "Selection",    AutoSize = true };
        private readonly RadioButton _rbSheet   = new RadioButton { Text = "Active sheet", AutoSize = true, Checked = true };
        private readonly RadioButton _rbAll     = new RadioButton { Text = "All sheets",   AutoSize = true };
        private readonly CheckBox    _chkCase   = new CheckBox   { Text = "Match case",    AutoSize = true };
        private readonly Label       _lblStatus = new Label { AutoSize = true, ForeColor = Color.DimGray };
        private readonly ListView    _grid      = new ListView
        {
            FullRowSelect = true, View = View.Details, GridLines = true,
            BorderStyle   = BorderStyle.FixedSingle, HideSelection = false
        };

        public RegexFindReplaceDialog()
        {
            Text       = "Regex Find & Replace";
            ClientSize = new Size(620, 420);

            var lblF = new Label { Text = "Find (regex):", Left = 12, Top = 14, AutoSize = true };
            var lblR = new Label { Text = "Replace with:", Left = 12, Top = 44, AutoSize = true };
            _find   .SetBounds(118, 11, 356, 23);
            _replace.SetBounds(118, 41, 356, 23);

            var tip = new Label
            {
                Text      = "Tip: use ( ) to capture groups, $1 $2 in replacement.  e.g.  find: (\\d+)-(\\w+)   replace: $2-$1",
                Left      = 12, Top = 70, Width = 590, AutoSize = false,
                ForeColor = Color.DimGray, Font = new Font("Segoe UI", 8f)
            };

            var btnPreview = new Button { Text = "&Preview",     Left = 482, Top = 10, Width = 92, Height = 23 };
            var btnReplace = new Button { Text = "Replace &All", Left = 482, Top = 40, Width = 92, Height = 23 };
            btnPreview.Click += OnPreview;
            btnReplace.Click += OnReplaceAll;

            var grpScope = new GroupBox { Text = "Scope", Left = 12, Top = 90, Width = 310, Height = 44 };
            _rbSel  .SetBounds(8,   18, 80, 18);
            _rbSheet.SetBounds(96,  18, 96, 18);
            _rbAll  .SetBounds(200, 18, 88, 18);
            grpScope.Controls.AddRange(new Control[] { _rbSel, _rbSheet, _rbAll });

            _chkCase.SetBounds(338, 104, 100, 18);
            _lblStatus.SetBounds(12, 142, 590, 18);

            _grid.SetBounds(12, 164, 590, 216);
            _grid.Columns.Add("Sheet",    100);
            _grid.Columns.Add("Cell",      58);
            _grid.Columns.Add("Original", 210);
            _grid.Columns.Add("→ Result", 192);

            var btnClose = new Button { Text = "&Close", Left = 526, Top = 388, Width = 82, DialogResult = DialogResult.Cancel };

            WireButtons(null, btnClose);
            Controls.AddRange(new Control[]
            {
                lblF, lblR, _find, _replace, tip,
                btnPreview, btnReplace,
                grpScope, _chkCase,
                _lblStatus, _grid, btnClose
            });
            ActiveControl = _find;
        }

        // ── Regex builder ─────────────────────────────────────────────────────

        private bool TryBuildRegex(out Regex rx)
        {
            rx = null;
            if (string.IsNullOrEmpty(_find.Text)) { SetError(_find, "Enter a pattern."); return false; }
            SetError(_find, null);
            try
            {
                var opts = _chkCase.Checked ? RegexOptions.None : RegexOptions.IgnoreCase;
                rx = new Regex(_find.Text, opts, TimeSpan.FromSeconds(5));
                return true;
            }
            catch (ArgumentException ex) { SetError(_find, "Invalid pattern: " + ex.Message); return false; }
        }

        // ── Core scan ─────────────────────────────────────────────────────────

        private struct MatchRow { public string Sheet, Address, Original, Result; }

        private List<MatchRow> Scan(Regex rx, bool apply)
        {
            var app  = Globals.ThisAddIn.Application;
            var rows = new List<MatchRow>();
            string repl = _replace.Text;

            var sheets = new List<Excel.Worksheet>();
            if (_rbAll.Checked)
                foreach (Excel.Worksheet ws in app.ActiveWorkbook.Worksheets) sheets.Add(ws);
            else if (app.ActiveSheet is Excel.Worksheet cur)
                sheets.Add(cur);

            try
            {
                if (apply) { app.ScreenUpdating = false; app.EnableEvents = false; }

                foreach (var ws in sheets)
                {
                    Excel.Range scope = _rbSel.Checked
                        ? app.Selection as Excel.Range
                        : ws.UsedRange;
                    if (scope == null) continue;

                    // Bulk-read values — one COM call per sheet, much faster than cell-by-cell.
                    object raw  = scope.Value2;
                    object[,] vals = raw as object[,];

                    if (vals == null)
                    {
                        // Single-cell range
                        string sv = raw?.ToString();
                        if (sv != null && rx.IsMatch(sv))
                        {
                            string r = rx.Replace(sv, repl);
                            rows.Add(new MatchRow { Sheet = ws.Name, Address = scope.Address, Original = sv, Result = r });
                            if (apply) scope.Value2 = r;
                        }
                        continue;
                    }

                    int rowBase = scope.Row;
                    int colBase = scope.Column;

                    for (int ri = 1; ri <= vals.GetLength(0); ri++)
                    {
                        for (int ci = 1; ci <= vals.GetLength(1); ci++)
                        {
                            string sv = vals[ri, ci]?.ToString();
                            if (sv == null || !rx.IsMatch(sv)) continue;
                            string r  = rx.Replace(sv, repl);
                            var cell  = ws.Cells[ri + rowBase - 1, ci + colBase - 1] as Excel.Range;
                            rows.Add(new MatchRow { Sheet = ws.Name, Address = cell?.Address ?? "", Original = sv, Result = r });
                            if (apply && cell != null) cell.Value2 = r;
                            if (rows.Count >= 1000) return rows;
                        }
                    }
                }
            }
            finally
            {
                if (apply) try { app.ScreenUpdating = true; app.EnableEvents = true; } catch { }
            }
            return rows;
        }

        // ── Event handlers ────────────────────────────────────────────────────

        private void OnPreview(object sender, EventArgs e)
        {
            if (!TryBuildRegex(out Regex rx)) return;
            var rows = Scan(rx, apply: false);
            LoadGrid(rows);
            _lblStatus.ForeColor = Color.DimGray;
            _lblStatus.Text = rows.Count == 0    ? "No matches found."
                            : rows.Count >= 1000 ? "First 1 000 matches shown — refine the pattern to see fewer."
                            : rows.Count + " match(es) found. Click Replace All to apply.";
        }

        private void OnReplaceAll(object sender, EventArgs e)
        {
            if (!TryBuildRegex(out Regex rx)) return;
            if (MessageBox.Show("Replace all matches? This cannot be undone via Ctrl+Z.",
                    Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;

            var rows = Scan(rx, apply: true);
            LoadGrid(rows);
            _lblStatus.ForeColor = rows.Count > 0 ? Color.DarkGreen : Color.DimGray;
            _lblStatus.Text = rows.Count + " cell(s) replaced.";
        }

        private void LoadGrid(List<MatchRow> rows)
        {
            _grid.Items.Clear();
            foreach (var m in rows)
            {
                var item = new ListViewItem(m.Sheet);
                item.SubItems.Add(m.Address);
                item.SubItems.Add(m.Original);
                item.SubItems.Add(m.Result);
                _grid.Items.Add(item);
            }
        }
    }
}
