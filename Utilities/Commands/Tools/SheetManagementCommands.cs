using System;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    [ExcelCommand]
    public sealed class UnhideAllSheetsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Sheet.UnhideAll", Label = "Unhide All Sheets",
            Screentip = "Unhide All Sheets",
            Supertip = "Make every hidden worksheet in the workbook visible in one click.",
            ImageId = "UnhideAllSheets", Tab = "Workbook & Sheets", Group = "Sheets", Order = 50,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            Excel.Workbook wb = ctx.App.ActiveWorkbook; if (wb == null) return;
            int count = 0;
            foreach (Excel.Worksheet ws in wb.Worksheets)
                if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible) { ws.Visible = Excel.XlSheetVisibility.xlSheetVisible; count++; }
            MessageBox.Show(count > 0 ? count + " sheet(s) unhidden." : "No hidden sheets found.",
                Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    [ExcelCommand]
    public sealed class BatchRenameSheetsCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Sheet.BatchRename", Label = "Batch Rename Sheets",
            Screentip = "Batch Rename Sheets",
            Supertip = "Add a prefix or suffix to every sheet name in the workbook at once.",
            ImageId = "BatchRenameSheets", Tab = "Workbook & Sheets", Group = "Sheets", Order = 15,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new BatchRenameSheetsDialog();
    }

    internal sealed class BatchRenameSheetsDialog : Dialogs.DialogBase
    {
        private readonly TextBox _prefix = new TextBox();
        private readonly TextBox _suffix = new TextBox();
        public BatchRenameSheetsDialog()
        {
            Text = BatchRenameSheetsCommand.Def.Label; ClientSize = new Size(320, 130);
            var lblP = new Label { Text = "Prefix:", Left = 12, Top = 18, AutoSize = true };
            _prefix.SetBounds(70, 15, 232, 23);
            var lblS = new Label { Text = "Suffix:", Left = 12, Top = 50, AutoSize = true };
            _suffix.SetBounds(70, 47, 232, 23);
            var apply = new Button { Text = "&Apply", Left = 126, Top = 90, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 212, Top = 90, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string pre = _prefix.Text, suf = _suffix.Text;
                if (pre.Length == 0 && suf.Length == 0) { SetError(_prefix, "Enter a prefix or suffix."); return; }
                SetError(_prefix, null);
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook; if (wb == null) { Close(); return; }
                int count = 0;
                foreach (Excel.Worksheet ws in wb.Worksheets) { try { ws.Name = pre + ws.Name + suf; count++; } catch { } }
                Close();
                MessageBox.Show(count + " sheet(s) renamed.", BatchRenameSheetsCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            };
            Controls.AddRange(new Control[] { lblP, _prefix, lblS, _suffix, apply, cancel });
            WireButtons(apply, cancel); ActiveControl = _prefix;
        }
    }

    [ExcelCommand]
    public sealed class RefreshAllPivotsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Sheet.RefreshPivots", Label = "Refresh All Pivots",
            Screentip = "Refresh All Pivot Tables",
            Supertip = "Refresh every pivot table and pivot chart in the workbook from its source data.",
            ImageId = "RefreshPivots", Tab = "Workbook & Sheets", Group = "Workbook", Order = 30,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            Excel.Workbook wb = ctx.App.ActiveWorkbook; if (wb == null) return;
            int count = 0;
            foreach (Excel.Worksheet ws in wb.Worksheets)
                foreach (Excel.PivotTable pt in ws.PivotTables()) { pt.RefreshTable(); count++; }
            MessageBox.Show(count > 0 ? count + " pivot table(s) refreshed." : "No pivot tables found.",
                Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    [ExcelCommand]
    public sealed class ClearHyperlinksCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.ClearHyperlinks", Label = "Clear Hyperlinks",
            Screentip = "Clear All Hyperlinks",
            Supertip = "Remove every hyperlink from the selection, keeping the displayed text intact.",
            ImageId = "ClearHyperlinks", Tab = "Editing", Group = "Format & Convert", Order = 50,
            Scope = CommandScope.Selection, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            ctx.Target.Hyperlinks.Delete();
            MessageBox.Show("Hyperlinks cleared.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    [ExcelCommand]
    public sealed class AutoFitAllCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.AutoFitAll", Label = "AutoFit All",
            Screentip = "AutoFit All Columns and Rows",
            Supertip = "Automatically resize every column and row on the active sheet to fit its contents.",
            ImageId = "AutoFitAll", Tab = "Editing", Group = "Format & Convert", Order = 55,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet; if (ws == null) return;
            ws.Cells.EntireColumn.AutoFit(); ws.Cells.EntireRow.AutoFit();
            MessageBox.Show("All columns and rows auto-fitted.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    [ExcelCommand]
    public sealed class SelectFirstCellCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.FirstCell", Label = "Select First Cell",
            Screentip = "Go to First Data Cell",
            Supertip = "Navigate to the first cell that contains data on the active sheet.",
            ImageId = "SelectFirstCell", Tab = "Select & Navigate", Group = "Navigate", Order = 60,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet; if (ws == null) return;
            try { ws.UsedRange.Cells[1, 1].Select(); }
            catch { MessageBox.Show("No data found on this sheet.", Definition.Label); }
        }
    }

    [ExcelCommand]
    public sealed class SelectLastCellCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.LastCell", Label = "Select Last Cell",
            Screentip = "Go to Last Data Cell",
            Supertip = "Navigate to the last cell that contains data on the active sheet (bottom-right of the used range).",
            ImageId = "SelectLastCell", Tab = "Select & Navigate", Group = "Navigate", Order = 61,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet; if (ws == null) return;
            try { Excel.Range u = ws.UsedRange; u.Cells[u.Rows.Count, u.Columns.Count].Select(); }
            catch { MessageBox.Show("No data found on this sheet.", Definition.Label); }
        }
    }

    [ExcelCommand]
    public sealed class FreezePanesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.FreezePanes", Label = "Freeze Panes",
            Screentip = "Freeze / Unfreeze Panes",
            Supertip = "Freeze panes at the active cell, or unfreeze if already frozen.",
            ImageId = "FreezePanes", Tab = "Workbook & Sheets", Group = "View", Order = 10,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            Excel.Window win = ctx.App.ActiveWindow; if (win == null) return;
            if (win.FreezePanes) { win.FreezePanes = false; MessageBox.Show("Panes unfrozen.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information); }
            else { win.FreezePanes = true; MessageBox.Show("Panes frozen at " + (ctx.App.ActiveCell != null ? ctx.App.ActiveCell.Address : "selected cell") + ".", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }
    }

    [ExcelCommand]
    public sealed class HighlightDuplicatesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Data.HighlightDuplicates", Label = "Highlight Duplicates",
            Screentip = "Highlight Duplicate Values",
            Supertip = "Mark cells with duplicate values in yellow and unique values in green for quick visual identification.",
            ImageId = "HighlightDuplicates", Tab = "Data & Cleaning", Group = "Duplicates", Order = 15,
            Scope = CommandScope.Selection, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            var counts = new System.Collections.Generic.Dictionary<string, int>(StringComparer.Ordinal);
            foreach (Excel.Range cell in ctx.Target.Cells)
            { object v = cell.Value2; if (v == null) continue; string k = v.ToString(); counts[k] = counts.ContainsKey(k) ? counts[k] + 1 : 1; }
            int dups = 0, uniq = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2; if (v == null) continue;
                int c; counts.TryGetValue(v.ToString(), out c);
                if (c > 1) { cell.Interior.Color = Excel.XlRgbColor.rgbLightYellow; dups++; }
                else { cell.Interior.Color = Excel.XlRgbColor.rgbLightGreen; uniq++; }
            }
            MessageBox.Show("Duplicates: " + dups + "  Uniques: " + uniq, Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
