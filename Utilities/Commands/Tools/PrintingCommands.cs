using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using utilities.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    // ── 1. Print Multiple Workbooks ───────────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintMultipleWorkbooksCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintMultipleWorkbooks", Label = "Print Multiple Workbooks Wizard...",
            Screentip = "Print Multiple Workbooks",
            Supertip = "Print all selected open workbooks in one operation.",
            ImageId = "PrintMultipleWorkbooks", Tab = "Printing", Group = "Printing", Order = 10,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new PrintMultipleWorkbooksDialog(); }
    }

    internal sealed class PrintMultipleWorkbooksDialog : DialogBase
    {
        private readonly CheckedListBox _list = new CheckedListBox { Left = 12, Top = 40, Width = 340, Height = 180 };

        public PrintMultipleWorkbooksDialog()
        {
            Text = "Print Multiple Workbooks";
            ClientSize = new Size(370, 280);

            var lbl    = new Label { Text = "Select workbooks to print:", Left = 12, Top = 14, AutoSize = true };
            var print  = new Button { Text = "&Print",  Left = 182, Top = 238, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 270, Top = 238, Width = 80, DialogResult = DialogResult.Cancel };
            print.Click += OnPrint;
            Controls.AddRange(new Control[] { lbl, _list, print, cancel });
            WireButtons(print, cancel);

            foreach (Excel.Workbook wb in Globals.ThisAddIn.Application.Workbooks)
                _list.Items.Add(wb.Name, true);
        }

        private void OnPrint(object sender, EventArgs e)
        {
            var names = _list.CheckedItems.Cast<string>().ToList();
            if (names.Count == 0) { SetError(_list, "Select at least one workbook."); return; }
            SetError(_list, null);

            foreach (Excel.Workbook wb in Globals.ThisAddIn.Application.Workbooks)
                if (names.Contains(wb.Name)) wb.PrintOut();
            Close();
        }
    }

    // ── 2. Print Multiple Selections ─────────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintMultipleSelectionsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintMultipleSelections", Label = "Print Multiple Selections Wizard...",
            Screentip = "Print Multiple Selections",
            Supertip = "Combine several non-contiguous ranges onto a single print job.",
            ImageId = "PrintMultipleSelections", Tab = "Printing", Group = "Printing", Order = 20,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new PrintMultipleSelectionsDialog(); }
    }

    internal sealed class PrintMultipleSelectionsDialog : DialogBase
    {
        private readonly ListBox _ranges = new ListBox { Left = 12, Top = 40, Width = 260, Height = 120 };
        private readonly TextBox _addr   = new TextBox { Left = 12, Top = 172, Width = 200 };

        public PrintMultipleSelectionsDialog()
        {
            Text = "Print Multiple Selections";
            ClientSize = new Size(370, 260);

            var lbl1   = new Label { Text = "Ranges to print:", Left = 12, Top = 14, AutoSize = true };
            var btnAdd = new Button { Text = "Add",    Left = 282, Top = 40,  Width = 76 };
            var btnDel = new Button { Text = "Remove", Left = 282, Top = 70,  Width = 76 };
            var lbl2   = new Label { Text = "Range address:", Left = 12, Top = 148, AutoSize = true };
            var btnSel = new Button { Text = "Use Selection", Left = 220, Top = 170, Width = 138 };
            var print  = new Button { Text = "&Print",  Left = 182, Top = 220, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 270, Top = 220, Width = 80, DialogResult = DialogResult.Cancel };

            btnAdd.Click += (s, ev) => { if (_addr.Text.Length > 0) { _ranges.Items.Add(_addr.Text); _addr.Clear(); } };
            btnDel.Click += (s, ev) => { if (_ranges.SelectedIndex >= 0) _ranges.Items.RemoveAt(_ranges.SelectedIndex); };
            btnSel.Click += (s, ev) => { var r = Globals.ThisAddIn.Application.Selection as Excel.Range; if (r != null) _addr.Text = r.Address; };
            print.Click  += OnPrint;

            Controls.AddRange(new Control[] { lbl1, _ranges, btnAdd, btnDel, lbl2, _addr, btnSel, print, cancel });
            WireButtons(print, cancel);
        }

        private void OnPrint(object sender, EventArgs e)
        {
            if (_ranges.Items.Count == 0) { SetError(_ranges, "Add at least one range."); return; }
            SetError(_ranges, null);

            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;

            string combined = string.Join(",", _ranges.Items.Cast<string>());
            string prev = ws.PageSetup.PrintArea;
            ws.PageSetup.PrintArea = combined;
            ws.PrintOut();
            ws.PageSetup.PrintArea = prev;
            Close();
        }
    }

    // ── 3. Print First Page of Each Worksheet ────────────────────────────────

    [ExcelCommand]
    public sealed class PrintFirstPageCommand : CommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintFirstPage", Label = "Print First Page of Each Worksheet",
            Screentip = "Print First Page",
            Supertip = "Print page 1 of every worksheet in the active workbook.",
            ImageId = "PrintFirstPage", Tab = "Printing", Group = "Printing", Order = 30,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override void Run(CommandContext ctx)
        {
            foreach (Excel.Worksheet ws in ctx.App.ActiveWorkbook.Worksheets)
                ws.PrintOut(From: 1, To: 1);
        }
    }

    // ── 4. Print Pages in Reverse Order ──────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintReverseOrderCommand : CommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintReverseOrder", Label = "Print Pages in Reverse Order",
            Screentip = "Print Reverse Order",
            Supertip = "Print the active worksheet from the last page to the first.",
            ImageId = "PrintReverseOrder", Tab = "Printing", Group = "Printing", Order = 40,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            int pages = ws.PageSetup.Pages.Count;
            for (int p = pages; p >= 1; p--)
                ws.PrintOut(From: p, To: p);
        }
    }

    // ── 5. Print Current Page ─────────────────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintCurrentPageCommand : CommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintCurrentPage", Label = "Print Current Page",
            Screentip = "Print Current Page",
            Supertip = "Print only the page that contains the active cell.",
            ImageId = "PrintCurrentPage", Tab = "Printing", Group = "Printing", Order = 50,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            int pageNum = 1;
            try { pageNum = (int)(double)ctx.App.ExecuteExcel4Macro("GET.DOCUMENT(64)"); }
            catch { pageNum = 1; }
            ws.PrintOut(From: pageNum, To: pageNum);
        }
    }

    // ── 6. Print Specified Pages ──────────────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintSpecifiedPagesCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintSpecifiedPages", Label = "Print Specified Pages...",
            Screentip = "Print Specified Pages",
            Supertip = "Enter a page range (e.g. 1-3, 5) and print only those pages.",
            ImageId = "PrintSpecifiedPages", Tab = "Printing", Group = "Printing", Order = 60,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new PrintSpecifiedPagesDialog(); }
    }

    internal sealed class PrintSpecifiedPagesDialog : DialogBase
    {
        private readonly TextBox _pages = new TextBox { Left = 140, Top = 20, Width = 160 };

        public PrintSpecifiedPagesDialog()
        {
            Text = "Print Specified Pages";
            ClientSize = new Size(320, 100);

            var lbl    = new Label { Text = "Pages (e.g. 1-3,5):", Left = 12, Top = 22, AutoSize = true };
            var print  = new Button { Text = "&Print",  Left = 136, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 224, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            print.Click += OnPrint;
            Controls.AddRange(new Control[] { lbl, _pages, print, cancel });
            WireButtons(print, cancel);
            ActiveControl = _pages;
        }

        private void OnPrint(object sender, EventArgs e)
        {
            string raw = _pages.Text.Trim();
            if (raw.Length == 0) { SetError(_pages, "Enter page numbers."); return; }

            var pageNums = new List<int>();
            foreach (string part in raw.Split(','))
            {
                string p = part.Trim();
                if (p.Contains("-"))
                {
                    string[] bounds = p.Split('-');
                    int from, to;
                    if (bounds.Length == 2 && int.TryParse(bounds[0].Trim(), out from) && int.TryParse(bounds[1].Trim(), out to))
                        for (int i = from; i <= to; i++) pageNums.Add(i);
                }
                else { int n; if (int.TryParse(p, out n)) pageNums.Add(n); }
            }

            if (pageNums.Count == 0) { SetError(_pages, "No valid page numbers found."); return; }
            SetError(_pages, null);

            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            pageNums.Sort();
            foreach (int pg in pageNums.Distinct())
                ws.PrintOut(From: pg, To: pg);
            Close();
        }
    }

    // ── 7. Print Circle Invalid Data ─────────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintCircleInvalidDataCommand : CommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintCircleInvalidData", Label = "Print Circle Invalid Data ...",
            Screentip = "Print Circle Invalid Data",
            Supertip = "Circle all cells that fail data-validation rules, then print the sheet.",
            ImageId = "PrintCircleInvalidData", Tab = "Printing", Group = "Printing", Order = 70,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override void Run(CommandContext ctx)
        {
            dynamic ws = ctx.App.ActiveSheet;
            if (ws == null) return;
            ws.CircleInvalidData();
            ws.PrintOut();
            ws.ClearCircles();
        }
    }

    // ── 8. Print Charts Only ─────────────────────────────────────────────────

    [ExcelCommand]
    public sealed class PrintChartsOnlyCommand : CommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PrintChartsOnly", Label = "Print Charts Only...",
            Screentip = "Print Charts Only",
            Supertip = "Print every embedded chart in the active worksheet.",
            ImageId = "PrintChartsOnly", Tab = "Printing", Group = "Printing", Order = 80,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            int count = 0;
            foreach (Excel.ChartObject co in (Excel.ChartObjects)ws.ChartObjects())
            {
                co.Chart.PrintOut();
                count++;
            }
            if (count == 0)
                MessageBox.Show("No charts found on the active worksheet.", Def.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    // ── 9. Copy Page Setup ────────────────────────────────────────────────────

    [ExcelCommand]
    public sealed class CopyPageSetupCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.CopyPageSetup", Label = "Copy Page Setup...",
            Screentip = "Copy Page Setup",
            Supertip = "Copy page-setup settings from the active sheet to other sheets in the workbook.",
            ImageId = "CopyPageSetup", Tab = "Printing", Group = "Printing", Order = 90,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new CopyPageSetupDialog(); }
    }

    internal sealed class CopyPageSetupDialog : DialogBase
    {
        // ── Layout constants ──────────────────────────────────────────────────
        private const int LeftPanelW  = 180;
        private const int RightPanelX = 200;
        private const int RightPanelW = 260;
        private const int PanelH      = 390;

        private readonly CheckedListBox _sheets = new CheckedListBox();
        private readonly TreeView       _tree   = new TreeView();
        private bool _updating;

        // ── Option node keys ─────────────────────────────────────────────────
        private static readonly (string Key, string Label, bool Default)[] PageOpts =
        {
            ("Orientation",     "Orientation",         true),
            ("Zoom",            "Zoom",                false),
            ("FitPageWide",     "Fit page wide",       true),
            ("FitPageTall",     "Fit page tall",       false),
            ("PaperSize",       "Paper size",          true),
            ("PrintQuality",    "Print quality",       false),
            ("FirstPageNum",    "First page number",   false),
        };
        private static readonly (string Key, string Label, bool Default)[] MarginOpts =
        {
            ("LeftMargin",      "Left margin",                 true),
            ("RightMargin",     "Right margin",                true),
            ("TopMargin",       "Top margin",                  true),
            ("BottomMargin",    "Bottom margin",               true),
            ("HeaderMargin",    "Header margin",               true),
            ("FooterMargin",    "Footer margin",               true),
            ("CenterH",         "Center on page horizontally", true),
            ("CenterV",         "Center on page vertically",   true),
        };
        private static readonly (string Key, string Label, bool Default)[] HFOpts =
        {
            ("LeftHeader",      "Left header",                false),
            ("CenterHeader",    "Center header",              false),
            ("RightHeader",     "Right header",               false),
            ("PicHeaders",      "Pictures headers",           false),
            ("LeftFooter",      "Left footer",                true),
            ("CenterFooter",    "Center footer",              false),
            ("RightFooter",     "Right footer",               true),
            ("PicFooters",      "Pictures footers",           false),
            ("OddEvenHF",       "Different odd and even pages", false),
            ("DifferentFirst",  "Different first page",       false),
            ("ScaleWithDoc",    "Scale with document",        false),
            ("AlignMargins",    "Align with page margins",    false),
        };
        private static readonly (string Key, string Label, bool Default)[] SheetOpts =
        {
            ("PrintArea",       "Print area",              false),
            ("RepeatRows",      "Rows to repeat at top",   false),
            ("RepeatCols",      "Columns to repeat at left", false),
            ("Gridlines",       "Gridlines",               false),
            ("BlackWhite",      "Black and white",         false),
            ("DraftQuality",    "Draft quality",           false),
            ("RowColHeadings",  "Row and column headings", false),
            ("Comments",        "Comments",                false),
            ("CellErrors",      "Cell errors",             false),
            ("PrintOrder",      "Print Order",             false),
        };

        public CopyPageSetupDialog()
        {
            Text        = "Copy Page Setup";
            ClientSize  = new Size(480, 470);
            MinimumSize = new Size(480, 470);

            // ── Left panel ────────────────────────────────────────────────────
            var lblSheets = new Label { Text = "Copy to", Left = 12, Top = 8, AutoSize = true, Font = new Font(Font, FontStyle.Bold) };

            _sheets.SetBounds(12, 28, LeftPanelW, PanelH);
            _sheets.CheckOnClick = true;

            string activeName = (Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet)?.Name ?? "";
            foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                if (ws.Name != activeName) _sheets.Items.Add(ws.Name, true);

            // ── Right panel ───────────────────────────────────────────────────
            var lblOpts = new Label { Text = "Options", Left = RightPanelX, Top = 8, AutoSize = true, Font = new Font(Font, FontStyle.Bold) };

            _tree.SetBounds(RightPanelX, 28, RightPanelW, PanelH);
            _tree.CheckBoxes   = true;
            _tree.ShowLines    = true;
            _tree.ShowPlusMinus = true;
            _tree.AfterCheck  += OnAfterCheck;
            BuildTree();

            // ── Buttons ───────────────────────────────────────────────────────
            var ok     = new Button { Text = "OK",     Left = 296, Top = 430, Width = 80, DialogResult = DialogResult.None };
            var cancel = new Button { Text = "Cancel", Left = 386, Top = 430, Width = 80, DialogResult = DialogResult.Cancel };
            ok.Click += OnOk;

            Controls.AddRange(new Control[] { lblSheets, _sheets, lblOpts, _tree, ok, cancel });
            WireButtons(ok, cancel);
        }

        private void BuildTree()
        {
            _tree.BeginUpdate();
            var root = new TreeNode("Page setup") { Checked = true };

            root.Nodes.Add(BuildSection("Page",          PageOpts));
            root.Nodes.Add(BuildSection("Margins",       MarginOpts));
            root.Nodes.Add(BuildSection("Header/Footer", HFOpts));
            root.Nodes.Add(BuildSection("Sheet",         SheetOpts));

            _tree.Nodes.Add(root);
            root.Expand();
            foreach (TreeNode s in root.Nodes) s.Expand();
            UpdateParent(root);
            _tree.EndUpdate();
        }

        private static TreeNode BuildSection(string name, (string Key, string Label, bool Default)[] opts)
        {
            var section = new TreeNode(name);
            foreach (var o in opts)
                section.Nodes.Add(new TreeNode(o.Label) { Name = o.Key, Checked = o.Default });
            return section;
        }

        private void OnAfterCheck(object sender, TreeViewEventArgs e)
        {
            if (_updating || e.Action == TreeViewAction.Unknown) return;
            _updating = true;
            try
            {
                // Propagate down to children
                SetChildrenChecked(e.Node, e.Node.Checked);
                // Propagate up to parents
                var p = e.Node.Parent;
                while (p != null) { UpdateParent(p); p = p.Parent; }
            }
            finally { _updating = false; }
        }

        private static void SetChildrenChecked(TreeNode node, bool check)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = check;
                SetChildrenChecked(child, check);
            }
        }

        private static void UpdateParent(TreeNode parent)
        {
            if (parent.Nodes.Count == 0) return;
            int checkedCount = 0;
            foreach (TreeNode c in parent.Nodes) if (c.Checked) checkedCount++;
            parent.Checked = checkedCount > 0;
        }

        private bool IsChecked(string key)
        {
            return FindNode(_tree.Nodes, key)?.Checked ?? false;
        }

        private static TreeNode FindNode(TreeNodeCollection nodes, string key)
        {
            foreach (TreeNode n in nodes)
            {
                if (n.Name == key) return n;
                var found = FindNode(n.Nodes, key);
                if (found != null) return found;
            }
            return null;
        }

        private void OnOk(object sender, EventArgs e)
        {
            var targets = _sheets.CheckedItems.Cast<string>().ToList();
            if (targets.Count == 0) { SetError(_sheets, "Select at least one target sheet."); return; }
            SetError(_sheets, null);

            Excel.Worksheet src = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (src == null) return;
            Excel.PageSetup sp = src.PageSetup;

            foreach (Excel.Worksheet ws in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                if (!targets.Contains(ws.Name)) continue;
                Excel.PageSetup dp = ws.PageSetup;

                // Page
                if (IsChecked("Orientation"))  dp.Orientation     = sp.Orientation;
                if (IsChecked("Zoom"))          dp.Zoom            = sp.Zoom;
                if (IsChecked("FitPageWide"))   dp.FitToPagesWide  = sp.FitToPagesWide;
                if (IsChecked("FitPageTall"))   dp.FitToPagesTall  = sp.FitToPagesTall;
                if (IsChecked("PaperSize"))     dp.PaperSize       = sp.PaperSize;
                if (IsChecked("PrintQuality"))  dp.PrintQuality    = sp.PrintQuality;
                if (IsChecked("FirstPageNum"))  dp.FirstPageNumber = sp.FirstPageNumber;
                // Margins
                if (IsChecked("LeftMargin"))    dp.LeftMargin      = sp.LeftMargin;
                if (IsChecked("RightMargin"))   dp.RightMargin     = sp.RightMargin;
                if (IsChecked("TopMargin"))     dp.TopMargin       = sp.TopMargin;
                if (IsChecked("BottomMargin"))  dp.BottomMargin    = sp.BottomMargin;
                if (IsChecked("HeaderMargin"))  dp.HeaderMargin    = sp.HeaderMargin;
                if (IsChecked("FooterMargin"))  dp.FooterMargin    = sp.FooterMargin;
                if (IsChecked("CenterH"))       dp.CenterHorizontally = sp.CenterHorizontally;
                if (IsChecked("CenterV"))       dp.CenterVertically   = sp.CenterVertically;
                // Header/Footer
                if (IsChecked("LeftHeader"))    dp.LeftHeader      = sp.LeftHeader;
                if (IsChecked("CenterHeader"))  dp.CenterHeader    = sp.CenterHeader;
                if (IsChecked("RightHeader"))   dp.RightHeader     = sp.RightHeader;
                if (IsChecked("LeftFooter"))    dp.LeftFooter      = sp.LeftFooter;
                if (IsChecked("CenterFooter"))  dp.CenterFooter    = sp.CenterFooter;
                if (IsChecked("RightFooter"))   dp.RightFooter     = sp.RightFooter;
                if (IsChecked("OddEvenHF"))     dp.OddAndEvenPagesHeaderFooter    = sp.OddAndEvenPagesHeaderFooter;
                if (IsChecked("DifferentFirst")) dp.DifferentFirstPageHeaderFooter = sp.DifferentFirstPageHeaderFooter;
                if (IsChecked("ScaleWithDoc"))  dp.ScaleWithDocHeaderFooter       = sp.ScaleWithDocHeaderFooter;
                if (IsChecked("AlignMargins"))  dp.AlignMarginsHeaderFooter       = sp.AlignMarginsHeaderFooter;
                // Sheet
                if (IsChecked("PrintArea"))     dp.PrintArea       = sp.PrintArea;
                if (IsChecked("RepeatRows"))    dp.PrintTitleRows  = sp.PrintTitleRows;
                if (IsChecked("RepeatCols"))    dp.PrintTitleColumns = sp.PrintTitleColumns;
                if (IsChecked("Gridlines"))     dp.PrintGridlines  = sp.PrintGridlines;
                if (IsChecked("BlackWhite"))    dp.BlackAndWhite   = sp.BlackAndWhite;
                if (IsChecked("DraftQuality"))  dp.Draft           = sp.Draft;
                if (IsChecked("RowColHeadings")) dp.PrintHeadings  = sp.PrintHeadings;
                if (IsChecked("Comments"))      dp.PrintComments   = sp.PrintComments;
                if (IsChecked("CellErrors"))    dp.PrintErrors     = sp.PrintErrors;
                if (IsChecked("PrintOrder"))    dp.Order           = sp.Order;
            }
            Close();
        }
    }

    // ── 10. Paging Subtotals ──────────────────────────────────────────────────

    [ExcelCommand]
    public sealed class PagingSubtotalsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.PagingSubtotals", Label = "Paging Subtotals...",
            Screentip = "Paging Subtotals",
            Supertip = "Insert a page break after each subtotal row so every group prints on its own page.",
            ImageId = "PagingSubtotals", Tab = "Printing", Group = "Printing", Order = 100,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new PagingSubtotalsDialog(); }
    }

    internal sealed class PagingSubtotalsDialog : DialogBase
    {
        private readonly TextBox _keyword = new TextBox { Left = 160, Top = 20, Width = 160, Text = "Total" };

        public PagingSubtotalsDialog()
        {
            Text = "Paging Subtotals";
            ClientSize = new Size(340, 100);

            var lbl    = new Label { Text = "Subtotal row keyword:", Left = 12, Top = 22, AutoSize = true };
            var apply  = new Button { Text = "&Apply",  Left = 152, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 242, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnApply;
            Controls.AddRange(new Control[] { lbl, _keyword, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _keyword;
        }

        private void OnApply(object sender, EventArgs e)
        {
            string keyword = _keyword.Text.Trim();
            if (keyword.Length == 0) { SetError(_keyword, "Enter a keyword."); return; }
            SetError(_keyword, null);

            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Excel.Range used = ws.UsedRange;
            int lastRow = used.Row + used.Rows.Count - 1;
            int count = 0;

            for (int r = used.Row; r <= lastRow; r++)
            {
                string text = ((Excel.Range)ws.Cells[r, used.Column]).Text?.ToString() ?? "";
                if (text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (r < lastRow)
                    {
                        ((Excel.Range)ws.Rows[r + 1]).PageBreak = (int)Excel.XlPageBreak.xlPageBreakManual;
                        count++;
                    }
                }
            }

            MessageBox.Show(count + " page break(s) inserted.", "Paging Subtotals",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }
    }

    // ── 11. Insert Page Break Every Row ──────────────────────────────────────

    [ExcelCommand]
    public sealed class InsertPageBreakEveryRowCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.InsertPageBreakEveryRow", Label = "Insert Page Break Every Row...",
            Screentip = "Insert Page Break Every Row",
            Supertip = "Insert a horizontal page break after every N rows in the used range.",
            ImageId = "InsertPageBreakEveryRow", Tab = "Printing", Group = "Printing", Order = 110,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new InsertPageBreakEveryRowDialog(); }
    }

    internal sealed class InsertPageBreakEveryRowDialog : DialogBase
    {
        private readonly NumericUpDown _rows = new NumericUpDown
            { Left = 160, Top = 20, Width = 70, Minimum = 1, Maximum = 9999, Value = 50 };

        public InsertPageBreakEveryRowDialog()
        {
            Text = "Insert Page Break Every Row";
            ClientSize = new Size(310, 100);

            var lbl    = new Label { Text = "Insert break every:", Left = 12, Top = 22, AutoSize = true };
            var lbl2   = new Label { Text = "rows",  Left = 238, Top = 22, AutoSize = true };
            var apply  = new Button { Text = "&Apply",  Left = 122, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 214, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnApply;
            Controls.AddRange(new Control[] { lbl, _rows, lbl2, apply, cancel });
            WireButtons(apply, cancel);
        }

        private void OnApply(object sender, EventArgs e)
        {
            int n = (int)_rows.Value;
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Excel.Range used = ws.UsedRange;
            int startRow = used.Row;
            int lastRow  = startRow + used.Rows.Count - 1;

            ws.ResetAllPageBreaks();
            for (int r = startRow + n; r <= lastRow; r += n)
                ((Excel.Range)ws.Rows[r]).PageBreak = (int)Excel.XlPageBreak.xlPageBreakManual;

            Close();
        }
    }

    // ── 12. Add Border to Each Page ───────────────────────────────────────────

    [ExcelCommand]
    public sealed class AddBorderToEachPageCommand : CommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Printing.AddBorderToEachPage", Label = "Add Border to Each Page",
            Screentip = "Add Border to Each Page",
            Supertip = "Draw a thin outer border around the cell range covered by each printed page.",
            ImageId = "AddBorderToEachPage", Tab = "Printing", Group = "Printing", Order = 120,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Excel.Range used = ws.UsedRange;
            int startRow = used.Row, startCol = used.Column;
            int lastRow  = startRow + used.Rows.Count - 1;
            int lastCol  = startCol + used.Columns.Count - 1;

            var vBreaks = new List<int> { startCol };
            foreach (Excel.VPageBreak vb in ws.VPageBreaks)
                vBreaks.Add(vb.Location.Column);
            vBreaks.Add(lastCol + 1);

            var hBreaks = new List<int> { startRow };
            foreach (Excel.HPageBreak hb in ws.HPageBreaks)
                hBreaks.Add(hb.Location.Row);
            hBreaks.Add(lastRow + 1);

            vBreaks.Sort(); hBreaks.Sort();

            for (int hi = 0; hi < hBreaks.Count - 1; hi++)
            {
                for (int vi = 0; vi < vBreaks.Count - 1; vi++)
                {
                    int rFrom = hBreaks[hi], rTo = hBreaks[hi + 1] - 1;
                    int cFrom = vBreaks[vi], cTo = vBreaks[vi + 1] - 1;
                    if (rTo < rFrom || cTo < cFrom) continue;
                    ApplyOutlineBorder(ws.Range[ws.Cells[rFrom, cFrom], ws.Cells[rTo, cTo]]);
                }
            }
        }

        private static void ApplyOutlineBorder(Excel.Range r)
        {
            foreach (Excel.XlBordersIndex idx in new[]
            {
                Excel.XlBordersIndex.xlEdgeLeft,  Excel.XlBordersIndex.xlEdgeRight,
                Excel.XlBordersIndex.xlEdgeTop,   Excel.XlBordersIndex.xlEdgeBottom
            })
            {
                Excel.Border b = r.Borders[idx];
                b.LineStyle = Excel.XlLineStyle.xlContinuous;
                b.Weight    = Excel.XlBorderWeight.xlThin;
            }
        }
    }
}
