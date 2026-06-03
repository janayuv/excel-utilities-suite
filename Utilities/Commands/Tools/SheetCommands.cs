using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Create a new sheet listing every sheet name as a hyperlink (table-of-contents).</summary>
    [ExcelCommand]
    public sealed class CreateSheetTocCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Sheet.CreateTOC",
            Label = "Sheet TOC",
            Screentip = "Create Sheet Table of Contents",
            Supertip = "Insert a new sheet at the front of the workbook with a clickable hyperlink to every other sheet.",
            ImageId = "SheetTOC",
            Tab = "Workbook & Sheets",
            Group = "Sheets",
            Order = 30,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Workbook wb = ctx.App.ActiveWorkbook;
            if (wb == null) return;

            const string tocName = "Sheet Index";
            foreach (Excel.Worksheet s in wb.Worksheets)
            {
                if (s.Name == tocName) { s.Delete(); break; }
            }

            Excel.Worksheet toc = (Excel.Worksheet)wb.Sheets.Add(wb.Sheets[1]);
            toc.Name = tocName;
            ((Excel.Range)toc.Cells[1, 1]).Value2 = "Sheet Index";
            ((Excel.Range)toc.Cells[1, 1]).Font.Bold = true;

            int row = 2;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                if (ws.Name == tocName) continue;
                toc.Hyperlinks.Add(toc.Cells[row, 1], "", ws.Name + "!A1", ws.Name, ws.Name);
                row++;
            }

            toc.Columns[1].AutoFit();
            MessageBox.Show("Sheet index created on \"" + tocName + "\".", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Sort all sheets in the workbook alphabetically.</summary>
    [ExcelCommand]
    public sealed class SortSheetsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Sheet.Sort",
            Label = "Sort Sheets",
            Screentip = "Sort Sheets",
            Supertip = "Rearrange all sheets in the workbook alphabetically — ascending or descending.",
            ImageId = "SortSheets",
            Tab = "Workbook & Sheets",
            Group = "Sheets",
            Order = 20,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SortSheetsDialog();
    }

    internal sealed class SortSheetsDialog : Dialogs.DialogBase
    {
        private readonly RadioButton _asc = new RadioButton { Text = "A to Z (ascending)", Checked = true };
        private readonly RadioButton _desc = new RadioButton { Text = "Z to A (descending)" };

        public SortSheetsDialog()
        {
            Text = SortSheetsCommand.Def.Label;
            ClientSize = new System.Drawing.Size(280, 130);
            _asc.SetBounds(16, 16, 240, 22);
            _desc.SetBounds(16, 42, 240, 22);

            var apply = new Button { Text = "&Sort", Left = 90, Top = 88, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 178, Top = 88, Width = 80, DialogResult = System.Windows.Forms.DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                bool asc = _asc.Checked;
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (wb == null) { Close(); return; }
                var names = new List<string>();
                foreach (Excel.Worksheet ws in wb.Worksheets) names.Add(ws.Name);
                names.Sort(StringComparer.OrdinalIgnoreCase);
                if (!asc) names.Reverse();
                for (int i = 0; i < names.Count; i++)
                    wb.Sheets[names[i]].Move(wb.Sheets[i + 1]);
                Close();
                MessageBox.Show("Sheets sorted.", SortSheetsCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Controls.AddRange(new System.Windows.Forms.Control[] { _asc, _desc, apply, cancel });
            WireButtons(apply, cancel);
        }
    }

    /// <summary>Rename multiple sheets by find-and-replace on sheet names.</summary>
    [ExcelCommand]
    public sealed class RenameMultipleSheetsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Sheet.RenameMultiple",
            Label = "Rename Sheets",
            Screentip = "Rename Multiple Sheets",
            Supertip = "Find and replace text in sheet names across the whole workbook at once.",
            ImageId = "RenameSheets",
            Tab = "Workbook & Sheets",
            Group = "Sheets",
            Order = 10,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new RenameSheetsDialog();
    }

    internal sealed class RenameSheetsDialog : Dialogs.DialogBase
    {
        private readonly TextBox _find = new TextBox();
        private readonly TextBox _replace = new TextBox();

        public RenameSheetsDialog()
        {
            Text = RenameMultipleSheetsCommand.Def.Label;
            ClientSize = new System.Drawing.Size(340, 130);

            var lblF = new Label { Text = "Find:", Left = 12, Top = 18, AutoSize = true };
            _find.SetBounds(80, 15, 240, 23);
            var lblR = new Label { Text = "Replace:", Left = 12, Top = 50, AutoSize = true };
            _replace.SetBounds(80, 47, 240, 23);

            var apply = new Button { Text = "&Rename", Left = 146, Top = 88, Width = 84 };
            var cancel = new Button { Text = "&Cancel", Left = 238, Top = 88, Width = 80, DialogResult = System.Windows.Forms.DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                if (_find.TextLength == 0) { SetError(_find, "Enter text to find."); return; }
                SetError(_find, null);
                string find = _find.Text, rep = _replace.Text;
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (wb == null) { Close(); return; }
                int count = 0;
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    if (ws.Name.Contains(find))
                    {
                        try { ws.Name = ws.Name.Replace(find, rep); count++; }
                        catch { }
                    }
                }
                Close();
                MessageBox.Show(count + " sheet(s) renamed.", RenameMultipleSheetsCommand.Def.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            Controls.AddRange(new System.Windows.Forms.Control[] { lblF, _find, lblR, _replace, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _find;
        }
    }

    /// <summary>Export each sheet in the workbook to a separate .xlsx file.</summary>
    [ExcelCommand]
    public sealed class ExportSheetsToFilesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Export.SheetsToFiles",
            Label = "Sheets to Files",
            Screentip = "Export Sheets to Separate Files",
            Supertip = "Save every sheet in the workbook as its own .xlsx file in a folder you choose.",
            ImageId = "SheetsToFiles",
            Tab = "Export / Import",
            Group = "Export",
            Order = 20,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Workbook wb = ctx.App.ActiveWorkbook;
            if (wb == null) return;

            using (var dlg = new FolderBrowserDialog { Description = "Choose output folder" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                string folder = dlg.SelectedPath;
                int exported = 0;
                int total = wb.Sheets.Count;

                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    string safeName = new string(ws.Name.Select(c =>
                        Path.GetInvalidFileNameChars().Contains(c) ? '_' : c).ToArray());
                    string path = Path.Combine(folder, safeName + ".xlsx");

                    ws.Copy();
                    Excel.Workbook newWb = ctx.App.ActiveWorkbook;
                    newWb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook,
                        Type.Missing, Type.Missing, false, false,
                        Excel.XlSaveAsAccessMode.xlNoChange,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    newWb.Close(false);
                    exported++;
                    ctx.Progress.Report((double)exported / total);
                }

                MessageBox.Show(exported + " sheet(s) exported to:\n" + folder,
                    Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
