using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using utilities.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    [ExcelCommand]
    public sealed class ExportToPdfCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Export.ToPdf",
            Label = "Export to PDF",
            Screentip = "Export to PDF",
            Supertip = "Export the selected range or active sheet to a PDF file. Choose quality, scaling, and print-area options before saving.",
            ImageId = "ExportToPDF",
            Tab = "Export / Import",
            Group = "Export",
            Order = 10,
            Scope = CommandScope.Selection,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Application app   = ctx.App;
            Excel.Range       sel   = app.Selection as Excel.Range;
            Excel.Worksheet   sheet = app.ActiveSheet as Excel.Worksheet;
            Excel.Workbook    wb    = app.ActiveWorkbook;

            bool hasSelection = sel != null && sel.Cells.Count > 1;
            string suggested  = BuildSuggestedName(wb, sheet);

            using (var dlg = new ExportToPdfDialog(hasSelection, suggested))
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                string path           = dlg.FilePath;
                bool   exportSel      = dlg.ExportSelection;
                var    quality        = dlg.UseMinimumSize
                                           ? Excel.XlFixedFormatQuality.xlQualityMinimum
                                           : Excel.XlFixedFormatQuality.xlQualityStandard;
                bool   ignorePrint    = !dlg.RespectPrintAreas;
                object source         = exportSel ? (object)sel : sheet;

                if (sheet != null && (dlg.FitToPageWide || dlg.Orientation != null))
                    ExportWithPageSetup(source, path, quality, ignorePrint, sheet,
                                        dlg.FitToPageWide, dlg.Orientation);
                else
                    DoExport(source, path, quality, ignorePrint);

                using (var post = new PostExportDialog(path))
                    post.ShowDialog();
            }
        }

        // Applies orientation and/or fit-to-one-page-wide, exports, then restores page setup.
        //
        // BUGFIX: late-bound PageSetup property writes intermittently throw
        // COMException 0x800A03EC ("Unable to set the FitToPagesTall property…").
        // The documented remedy (Excel 2010+) is to batch the changes by toggling
        // Application.PrintCommunication off while writing them, then back on.
        private static void ExportWithPageSetup(object source, string path,
            Excel.XlFixedFormatQuality quality, bool ignorePrintAreas, Excel.Worksheet sheet,
            bool fitWide, Excel.XlPageOrientation? orientation)
        {
            var app = Globals.ThisAddIn.Application;
            dynamic dps = sheet.PageSetup;

            // Capture originals for restore (best-effort).
            object origZoom = null, origOrient = null;
            int origFitWide = 1, origFitTall = 1;
            try { origZoom    = dps.Zoom; }                  catch { }
            try { origFitWide = (int)dps.FitToPagesWide; }   catch { }
            try { origFitTall = (int)dps.FitToPagesTall; }   catch { }
            try { origOrient  = dps.Orientation; }           catch { }

            bool prevComm = true;
            try { prevComm = app.PrintCommunication; app.PrintCommunication = false; } catch { }
            try
            {
                if (orientation != null) dps.Orientation = orientation.Value;
                if (fitWide)
                {
                    dps.Zoom           = false;
                    dps.FitToPagesWide = 1;
                    dps.FitToPagesTall = false;  // false = as many pages tall as needed
                }
            }
            finally { try { app.PrintCommunication = true; } catch { } }

            try { DoExport(source, path, quality, ignorePrintAreas); }
            finally
            {
                try { app.PrintCommunication = false; } catch { }
                try
                {
                    if (origOrient != null) dps.Orientation = origOrient;
                    if (fitWide)
                    {
                        if (origZoom is bool && !(bool)origZoom)
                        {
                            dps.Zoom           = false;
                            dps.FitToPagesWide = origFitWide;
                            dps.FitToPagesTall = origFitTall;
                        }
                        else
                        {
                            dps.Zoom = origZoom;
                        }
                    }
                }
                catch { }
                try { app.PrintCommunication = prevComm; } catch { }
            }
        }

        private static void DoExport(object source, string path,
            Excel.XlFixedFormatQuality quality, bool ignorePrintAreas)
        {
            ((dynamic)source).ExportAsFixedFormat(
                Excel.XlFixedFormatType.xlTypePDF,
                path,
                quality,
                true,             // include document properties
                ignorePrintAreas,
                Type.Missing, Type.Missing,
                false,            // do not open after publishing (we handle that ourselves)
                Type.Missing);
        }

        private static string BuildSuggestedName(Excel.Workbook wb, Excel.Worksheet ws)
        {
            string wbName = wb != null ? Path.GetFileNameWithoutExtension(wb.Name) : "Workbook";
            string wsName = ws != null ? ws.Name : "Sheet";
            string date   = DateTime.Today.ToString("yyyy-MM-dd");
            return SanitizeFileName(wbName) + " - " + SanitizeFileName(wsName) + " - " + date;
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }

    // ── Export options dialog ──────────────────────────────────────────────────

    internal sealed class ExportToPdfDialog : DialogBase
    {
        private readonly RadioButton _rbSelection;
        private readonly RadioButton _rbSheet;
        private readonly RadioButton _rbStandard;
        private readonly RadioButton _rbMinSize;
        private readonly CheckBox    _chkPrintAreas;
        private readonly CheckBox    _chkFitPage;
        private readonly ComboBox    _cmbOrient;
        private readonly TextBox     _pathBox;

        public string FilePath          { get; private set; }
        public bool   ExportSelection   => _rbSelection.Checked;
        public bool   UseMinimumSize    => _rbMinSize.Checked;
        public bool   RespectPrintAreas => _chkPrintAreas.Checked;
        public bool   FitToPageWide     => _chkFitPage.Checked;

        /// <summary>Null = keep the sheet's current orientation.</summary>
        public Excel.XlPageOrientation? Orientation
        {
            get
            {
                switch (_cmbOrient.SelectedIndex)
                {
                    case 1:  return Excel.XlPageOrientation.xlPortrait;
                    case 2:  return Excel.XlPageOrientation.xlLandscape;
                    default: return null;
                }
            }
        }

        public ExportToPdfDialog(bool hasSelection, string suggestedName)
        {
            Text       = "Export to PDF";
            ClientSize = new Size(440, 286);

            // ── Scope ──
            var grpScope = new GroupBox { Text = "Export scope", Left = 12, Top = 8, Width = 200, Height = 68 };
            _rbSelection = new RadioButton { Text = "Selection",   Left = 10, Top = 18, AutoSize = true, Enabled = hasSelection };
            _rbSheet     = new RadioButton { Text = "Active sheet", Left = 10, Top = 40, AutoSize = true };
            _rbSheet.Checked     = true;
            if (hasSelection) _rbSelection.Checked = true;
            grpScope.Controls.AddRange(new Control[] { _rbSelection, _rbSheet });

            // ── Quality ──
            var grpQuality = new GroupBox { Text = "Quality", Left = 224, Top = 8, Width = 204, Height = 68 };
            _rbStandard = new RadioButton { Text = "Standard (print)",            Left = 10, Top = 18, AutoSize = true, Checked = true };
            _rbMinSize  = new RadioButton { Text = "Minimum size (screen/email)", Left = 10, Top = 40, AutoSize = true };
            grpQuality.Controls.AddRange(new Control[] { _rbStandard, _rbMinSize });

            // ── Options ──
            var grpOpts = new GroupBox { Text = "Options", Left = 12, Top = 84, Width = 416, Height = 68 };
            _chkPrintAreas = new CheckBox { Text = "Respect defined print areas",                           Left = 10, Top = 18, AutoSize = true };
            _chkFitPage    = new CheckBox { Text = "Fit content to one page wide (overrides page scaling)", Left = 10, Top = 40, AutoSize = true };
            grpOpts.Controls.AddRange(new Control[] { _chkPrintAreas, _chkFitPage });

            // ── Orientation ──
            var lblOrient = new Label    { Text = "Orientation:", Left = 12, Top = 160, AutoSize = true };
            _cmbOrient    = new ComboBox { Left = 92, Top = 157, Width = 150, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbOrient.Items.AddRange(new object[] { "Keep current", "Portrait", "Landscape" });
            _cmbOrient.SelectedIndex = 0;

            // ── File path ──
            var lblPath   = new Label  { Text = "Save as:", Left = 12, Top = 196, AutoSize = true };
            _pathBox      = new TextBox { Left = 12, Top = 214, Width = 340 };
            var btnBrowse = new Button  { Text = "Browse…", Left = 358, Top = 212, Width = 70 };
            btnBrowse.Click += OnBrowse;

            string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            _pathBox.Text = Path.Combine(docs, suggestedName + ".pdf");

            // ── Buttons ──
            var btnExport = new Button { Text = "&Export", Left = 248, Top = 250, Width = 80 };
            var btnCancel = new Button { Text = "&Cancel", Left = 340, Top = 250, Width = 88, DialogResult = DialogResult.Cancel };
            btnExport.Click += OnExport;

            WireButtons(btnExport, btnCancel);
            Controls.AddRange(new Control[] { grpScope, grpQuality, grpOpts, lblOrient, _cmbOrient, lblPath, _pathBox, btnBrowse, btnExport, btnCancel });
        }

        private void OnBrowse(object sender, EventArgs e)
        {
            string current = _pathBox.Text.Trim();
            string dir = string.Empty;
            try { dir = Path.GetDirectoryName(current); } catch { }
            if (string.IsNullOrEmpty(dir))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            using (var sfd = new SaveFileDialog
            {
                Title            = "Save PDF",
                Filter           = "PDF document (*.pdf)|*.pdf",
                FileName         = Path.GetFileName(current),
                InitialDirectory = dir
            })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                    _pathBox.Text = sfd.FileName;
            }
        }

        private void OnExport(object sender, EventArgs e)
        {
            string path = _pathBox.Text.Trim();
            if (string.IsNullOrEmpty(path))
            {
                SetError(_pathBox, "Choose a file path.");
                return;
            }
            SetError(_pathBox, null);

            if (!path.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                path += ".pdf";

            string dir = string.Empty;
            try { dir = Path.GetDirectoryName(path); } catch { }
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                SetError(_pathBox, "Folder does not exist.");
                return;
            }

            FilePath     = path;
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ── Post-export dialog ─────────────────────────────────────────────────────

    internal sealed class PostExportDialog : DialogBase
    {
        public PostExportDialog(string path)
        {
            Text       = "Export Complete";
            ClientSize = new Size(420, 120);

            var lbl = new Label
            {
                Text     = "PDF exported successfully:\n" + path,
                Left     = 12, Top = 12, Width = 396, Height = 50,
                AutoSize = false
            };

            var btnOpen  = new Button { Text = "&Open PDF",       Left = 12,  Top = 80, Width = 90 };
            var btnShow  = new Button { Text = "&Show in Folder",  Left = 110, Top = 80, Width = 110 };
            var btnClose = new Button { Text = "&Close",           Left = 328, Top = 80, Width = 80, DialogResult = DialogResult.OK };

            btnOpen.Click += (s, e) =>
            {
                try { Process.Start(path); } catch { }
                DialogResult = DialogResult.OK;
                Close();
            };
            btnShow.Click += (s, e) =>
            {
                try { Process.Start("explorer.exe", "/select,\"" + path + "\""); } catch { }
                DialogResult = DialogResult.OK;
                Close();
            };

            WireButtons(btnClose, btnClose);
            Controls.AddRange(new Control[] { lbl, btnOpen, btnShow, btnClose });
        }
    }
}
