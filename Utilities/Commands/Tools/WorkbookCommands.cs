using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Copy one or more sheets to a different open workbook.</summary>
    [ExcelCommand]
    public sealed class CopySheetsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Sheet.CopySheets",
            Label = "Copy Sheets",
            Screentip = "Copy Sheets to Another Workbook",
            Supertip = "Copy the active sheet (or all sheets) into another open workbook.",
            ImageId = "CopySheets",
            Tab = "Workbook & Sheets",
            Group = "Sheets",
            Order = 40,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new CopySheetsDialog();
    }

    internal sealed class CopySheetsDialog : Dialogs.DialogBase
    {
        private readonly ComboBox _destWb = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly CheckBox _copyAll = new CheckBox { Text = "Copy all sheets (not just active)", Checked = false };

        public CopySheetsDialog()
        {
            Text = CopySheetsCommand.Def.Label;
            ClientSize = new System.Drawing.Size(360, 130);

            var lbl = new Label { Text = "Destination:", Left = 12, Top = 15, AutoSize = true };
            _destWb.SetBounds(95, 12, 250, 23);

            var app = Globals.ThisAddIn.Application;
            string sourceWbName = app.ActiveWorkbook != null ? app.ActiveWorkbook.Name : "";
            foreach (Excel.Workbook wb in app.Workbooks)
                if (wb.Name != sourceWbName) _destWb.Items.Add(wb.Name);

            if (_destWb.Items.Count > 0) _destWb.SelectedIndex = 0;
            else _destWb.Items.Add("(no other workbooks open)");

            _copyAll.SetBounds(12, 46, 330, 22);

            var apply = new Button { Text = "&Copy", Left = 166, Top = 88, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 252, Top = 88, Width = 80, DialogResult = System.Windows.Forms.DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                if (_destWb.Items.Count == 0 || _destWb.Text.StartsWith("("))
                {
                    MessageBox.Show("Open a destination workbook first.", CopySheetsCommand.Def.Label);
                    return;
                }
                Excel.Workbook dest = null;
                foreach (Excel.Workbook wb in Globals.ThisAddIn.Application.Workbooks)
                    if (wb.Name == _destWb.Text) { dest = wb; break; }
                if (dest == null) { Close(); return; }

                Excel.Workbook src = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (_copyAll.Checked)
                {
                    foreach (Excel.Worksheet ws in src.Worksheets)
                        ws.Copy(Type.Missing, dest.Sheets[dest.Sheets.Count]);
                }
                else
                {
                    ((Excel.Worksheet)src.ActiveSheet).Copy(Type.Missing, dest.Sheets[dest.Sheets.Count]);
                }

                Close();
                MessageBox.Show("Sheet(s) copied to " + dest.Name + ".",
                    CopySheetsCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Controls.AddRange(new System.Windows.Forms.Control[] { lbl, _destWb, _copyAll, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Merge all .xlsx files in a folder into the active workbook as separate sheets.</summary>
    [ExcelCommand]
    public sealed class MergeWorkbooksCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Sheet.MergeWorkbooks",
            Label = "Merge Workbooks",
            Screentip = "Merge Workbooks from Folder",
            Supertip = "Import every sheet from each .xlsx file in a chosen folder into the active workbook as new sheets.",
            ImageId = "MergeWorkbooks",
            Tab = "Workbook & Sheets",
            Group = "Workbook",
            Order = 10,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Workbook destWb = ctx.App.ActiveWorkbook;
            if (destWb == null) return;

            using (var dlg = new FolderBrowserDialog { Description = "Select folder containing .xlsx files" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                string[] files = Directory.GetFiles(dlg.SelectedPath, "*.xlsx");
                if (files.Length == 0)
                {
                    MessageBox.Show("No .xlsx files found in that folder.", Definition.Label,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int sheets = 0;
                for (int fi = 0; fi < files.Length; fi++)
                {
                    ctx.Progress.Report((double)fi / files.Length, "Merging " + Path.GetFileName(files[fi]));
                    Excel.Workbook srcWb = ctx.App.Workbooks.Open(files[fi], false, true);
                    try
                    {
                        foreach (Excel.Worksheet ws in srcWb.Worksheets)
                        {
                            ws.Copy(Type.Missing, destWb.Sheets[destWb.Sheets.Count]);
                            sheets++;
                        }
                    }
                    finally { srcWb.Close(false); }
                }

                MessageBox.Show(sheets + " sheet(s) merged from " + files.Length + " file(s).",
                    Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>Split the active sheet into separate workbooks, one per unique value in a key column.</summary>
    [ExcelCommand]
    public sealed class SplitSheetByColumnCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Sheet.SplitByColumn",
            Label = "Split Sheet by Column",
            Screentip = "Split Sheet by Column Value",
            Supertip = "Create a separate workbook for each unique value in a column — e.g. split a sales sheet by region.",
            ImageId = "SplitByColumn",
            Tab = "Workbook & Sheets",
            Group = "Workbook",
            Order = 20,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SplitByColumnDialog();
    }

    internal sealed class SplitByColumnDialog : Dialogs.DialogBase
    {
        private readonly NumericUpDown _colNum = new NumericUpDown { Minimum = 1, Maximum = 1000, Value = 1 };
        private readonly CheckBox _header = new CheckBox { Text = "First row is a header", Checked = true };

        public SplitByColumnDialog()
        {
            Text = SplitSheetByColumnCommand.Def.Label;
            ClientSize = new System.Drawing.Size(300, 130);

            var lbl = new Label { Text = "Key column:", Left = 12, Top = 15, AutoSize = true };
            _colNum.SetBounds(100, 12, 80, 23);
            _header.SetBounds(12, 46, 270, 22);

            var apply = new Button { Text = "&Split", Left = 106, Top = 90, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 90, Width = 80, DialogResult = System.Windows.Forms.DialogResult.Cancel };
            apply.Click += OnSplit;

            Controls.AddRange(new System.Windows.Forms.Control[] { lbl, _colNum, _header, apply, cancel });
            WireButtons(apply, cancel);
        }

        private void OnSplit(object sender, EventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet ws = app.ActiveSheet as Excel.Worksheet;
            if (ws == null) { Close(); return; }

            int keyCol = (int)_colNum.Value;
            bool hasHeader = _header.Checked;
            Excel.Range used = ws.UsedRange;
            int rows = used.Rows.Count;

            using (var dlg = new FolderBrowserDialog { Description = "Output folder" })
            {
                if (dlg.ShowDialog() != System.Windows.Forms.DialogResult.OK) { Close(); return; }

                var groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
                int startRow = hasHeader ? 2 : 1;
                for (int r = startRow; r <= rows; r++)
                {
                    object v = ((Excel.Range)ws.Cells[used.Row + r - 1, used.Column + keyCol - 1]).Value2;
                    string key = v != null ? v.ToString() : "(blank)";
                    if (!groups.ContainsKey(key)) groups[key] = new List<int>();
                    groups[key].Add(used.Row + r - 1);
                }

                int saved = 0;
                foreach (var kvp in groups)
                {
                    string safeName = new string(kvp.Key.Select(c =>
                        Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
                    string path = Path.Combine(dlg.SelectedPath, safeName + ".xlsx");

                    Excel.Workbook newWb = app.Workbooks.Add();
                    Excel.Worksheet newWs = (Excel.Worksheet)newWb.Sheets[1];

                    if (hasHeader)
                        ((Excel.Range)ws.Rows[used.Row]).Copy(newWs.Rows[1]);

                    int destRow = hasHeader ? 2 : 1;
                    foreach (int srcRow in kvp.Value)
                    {
                        ((Excel.Range)ws.Rows[srcRow]).Copy(newWs.Rows[destRow]);
                        destRow++;
                    }

                    newWb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook,
                        Type.Missing, Type.Missing, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    newWb.Close(false);
                    saved++;
                }

                Close();
                MessageBox.Show(saved + " file(s) created in:\n" + dlg.SelectedPath,
                    SplitSheetByColumnCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>Export the selected range as a PNG image file.</summary>
    [ExcelCommand]
    public sealed class ExportRangeAsImageCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Export.RangeAsImage",
            Label = "Range as Image",
            Screentip = "Export Range as Image",
            Supertip = "Copy the selected range to the clipboard as a picture, then save it as a PNG file.",
            ImageId = "ExportRangeImage",
            Tab = "Export / Import",
            Group = "Export",
            Order = 40,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            using (var dlg = new SaveFileDialog
            {
                Title = "Save Range as Image",
                Filter = "PNG image (*.png)|*.png",
                FileName = "range_export.png"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                ctx.Target.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlBitmap);
                System.Drawing.Image img = System.Windows.Forms.Clipboard.GetImage();
                if (img == null)
                {
                    MessageBox.Show("Could not capture the range as an image.", Definition.Label,
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                img.Save(dlg.FileName, System.Drawing.Imaging.ImageFormat.Png);
                img.Dispose();

                MessageBox.Show("Saved to:\n" + dlg.FileName, Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
