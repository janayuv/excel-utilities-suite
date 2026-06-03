using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    [ExcelCommand]
    public sealed class InsertFileListCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Import.FileList", Label = "Insert File List",
            Screentip = "Insert File List from Folder",
            Supertip = "List every file name, size, date and full path from a chosen folder into the active sheet.",
            ImageId = "InsertFileList", Tab = "Export / Import", Group = "Import", Order = 10,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            using (var dlg = new FolderBrowserDialog { Description = "Select folder to list files from" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                string[] files = Directory.GetFiles(dlg.SelectedPath);
                if (files.Length == 0) { MessageBox.Show("No files found.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet; if (ws == null) return;
                Excel.Range start = ctx.App.ActiveCell ?? ws.Cells[1, 1];
                int sr = start.Row, sc = start.Column;
                ((Excel.Range)ws.Cells[sr, sc]).Value2 = "File Name";
                ((Excel.Range)ws.Cells[sr, sc + 1]).Value2 = "Size (KB)";
                ((Excel.Range)ws.Cells[sr, sc + 2]).Value2 = "Last Modified";
                ((Excel.Range)ws.Cells[sr, sc + 3]).Value2 = "Full Path";
                ws.Range[ws.Cells[sr, sc], ws.Cells[sr, sc + 3]].Font.Bold = true;
                for (int i = 0; i < files.Length; i++)
                {
                    var fi = new FileInfo(files[i]); int row = sr + 1 + i;
                    ((Excel.Range)ws.Cells[row, sc]).Value2 = fi.Name;
                    ((Excel.Range)ws.Cells[row, sc + 1]).Value2 = Math.Round(fi.Length / 1024.0, 2);
                    ((Excel.Range)ws.Cells[row, sc + 2]).Value2 = fi.LastWriteTime.ToOADate();
                    ((Excel.Range)ws.Cells[row, sc + 2]).NumberFormat = "dd/mm/yyyy hh:mm";
                    ((Excel.Range)ws.Cells[row, sc + 3]).Value2 = fi.FullName;
                    if ((i & 0xFF) == 0) ctx.Progress.Report((double)i / files.Length);
                }
                ws.Columns[sc].AutoFit();
                MessageBox.Show(files.Length + " file(s) listed.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    [ExcelCommand]
    public sealed class ImportFilenamesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Import.Filenames", Label = "Import Filenames",
            Screentip = "Import Filenames",
            Supertip = "Open a file-picker and insert each selected file's full path into consecutive cells.",
            ImageId = "ImportFilenames", Tab = "Export / Import", Group = "Import", Order = 20,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            using (var dlg = new OpenFileDialog { Title = "Select files", Multiselect = true, Filter = "All files (*.*)|*.*" })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet; if (ws == null) return;
                Excel.Range start = ctx.App.ActiveCell ?? ws.Cells[1, 1];
                for (int i = 0; i < dlg.FileNames.Length; i++)
                    ((Excel.Range)ws.Cells[start.Row + i, start.Column]).Value2 = dlg.FileNames[i];
                MessageBox.Show(dlg.FileNames.Length + " filename(s) inserted.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    [ExcelCommand]
    public sealed class CountByColorCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Formula.CountByColor", Label = "Count by Color",
            Screentip = "Count Cells by Fill Color",
            Supertip = "Count how many cells share each fill colour in the selection and show a summary.",
            ImageId = "CountByColor", Tab = "Formula & Statistics", Group = "Statistics", Order = 15,
            Scope = CommandScope.Selection, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            var tally = new System.Collections.Generic.Dictionary<double, int>();
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                double color; try { color = (double)cell.Interior.Color; } catch { continue; }
                tally[color] = tally.ContainsKey(color) ? tally[color] + 1 : 1;
            }
            var sb = new System.Text.StringBuilder("Color count summary:\n\n");
            foreach (var kv in tally)
            {
                int rgb = (int)kv.Key;
                sb.AppendLine("RGB(" + (rgb & 0xFF) + "," + (rgb >> 8 & 0xFF) + "," + (rgb >> 16 & 0xFF) + "): " + kv.Value + " cell(s)");
            }
            MessageBox.Show(sb.ToString(), Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    [ExcelCommand]
    public sealed class MergeCellsKeepDataCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Data.MergeCellsKeepData", Label = "Merge Cells (Keep Data)",
            Screentip = "Merge Cells Without Losing Data",
            Supertip = "Merge selected cells into one, concatenating all values with a separator instead of discarding them.",
            ImageId = "MergeCellsKeepData", Tab = "Data & Cleaning", Group = "Merge", Order = 10,
            Scope = CommandScope.Selection, UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new MergeCellsKeepDataDialog();
    }

    internal sealed class MergeCellsKeepDataDialog : Dialogs.DialogBase
    {
        private readonly TextBox _sep = new TextBox { Text = " " };
        public MergeCellsKeepDataDialog()
        {
            Text = MergeCellsKeepDataCommand.Def.Label; ClientSize = new System.Drawing.Size(300, 100);
            var lbl = new Label { Text = "Separator:", Left = 12, Top = 18, AutoSize = true };
            _sep.SetBounds(88, 15, 192, 23);
            var apply = new Button { Text = "&Merge", Left = 106, Top = 58, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 58, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                string sep = _sep.Text;
                bool ok = RunOperation(MergeCellsKeepDataCommand.Def, CurrentSelection, ctx =>
                {
                    var sb = new System.Text.StringBuilder();
                    foreach (Excel.Range cell in ctx.Target.Cells) { object v = cell.Value2; if (v == null) continue; if (sb.Length > 0) sb.Append(sep); sb.Append(v.ToString()); }
                    ctx.Target.Merge(); ctx.Target.Cells[1, 1].Value2 = sb.ToString(); ctx.Target.WrapText = false;
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lbl, _sep, apply, cancel }); WireButtons(apply, cancel); ActiveControl = _sep;
        }
    }

    [ExcelCommand]
    public sealed class ClearConditionalFormattingCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Format.ClearConditionalFormatting", Label = "Clear Cond. Formatting",
            Screentip = "Clear Conditional Formatting",
            Supertip = "Remove all conditional formatting rules from the selected cells.",
            ImageId = "ClearCondFormatting", Tab = "Editing", Group = "Format & Convert", Order = 52,
            Scope = CommandScope.Selection, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            ctx.Target.FormatConditions.Delete();
            MessageBox.Show("Conditional formatting cleared.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    [ExcelCommand]
    public sealed class CopyVisibleCellsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.CopyVisibleCells", Label = "Copy Visible Cells",
            Screentip = "Copy Visible Cells Only",
            Supertip = "Copy only the visible cells in the selection to the clipboard, skipping hidden rows and columns.",
            ImageId = "CopyVisibleCells", Tab = "Utilities", Group = "Range & Cells", Order = 80,
            Scope = CommandScope.Selection, UndoMode = UndoMode.None
        };
        protected override void Run(CommandContext ctx)
        {
            try { ctx.Target.SpecialCells(Excel.XlCellType.xlCellTypeVisible).Copy(); MessageBox.Show("Visible cells copied to clipboard.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information); }
            catch { MessageBox.Show("No visible cells found.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }
    }

    [ExcelCommand]
    public sealed class InsertDateSequenceCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Insert.DateSequence", Label = "Insert Date Sequence",
            Screentip = "Insert Date Sequence",
            Supertip = "Fill the selected range with a sequence of dates — daily, weekdays only, weekly, or monthly.",
            ImageId = "InsertDateSequence", Tab = "Insert", Group = "Insert", Order = 40,
            Scope = CommandScope.Selection, UndoMode = UndoMode.FullSnapshot
        };
        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new InsertDateSequenceDialog();
    }

    internal sealed class InsertDateSequenceDialog : Dialogs.DialogBase
    {
        private readonly DateTimePicker _start = new DateTimePicker { Format = DateTimePickerFormat.Short };
        private readonly ComboBox _step = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        public InsertDateSequenceDialog()
        {
            Text = InsertDateSequenceCommand.Def.Label; ClientSize = new System.Drawing.Size(300, 120);
            var lblS = new Label { Text = "Start date:", Left = 12, Top = 18, AutoSize = true };
            _start.SetBounds(90, 14, 192, 23);
            var lblT = new Label { Text = "Step:", Left = 12, Top = 52, AutoSize = true };
            _step.Items.AddRange(new object[] { "Daily", "Weekdays only", "Weekly", "Monthly" });
            _step.SelectedIndex = 0; _step.SetBounds(90, 49, 192, 23);
            var apply = new Button { Text = "&Insert", Left = 106, Top = 82, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 192, Top = 82, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += (s, e) =>
            {
                DateTime date = _start.Value.Date; string step = _step.Text;
                bool ok = RunOperation(InsertDateSequenceCommand.Def, CurrentSelection, ctx =>
                {
                    DateTime cur = date;
                    foreach (Excel.Range cell in ctx.Target.Cells)
                    {
                        cell.Value2 = cur.ToOADate(); cell.NumberFormat = "dd/mm/yyyy";
                        switch (step)
                        {
                            case "Daily": cur = cur.AddDays(1); break;
                            case "Weekdays only": cur = cur.AddDays(1); while (cur.DayOfWeek == DayOfWeek.Saturday || cur.DayOfWeek == DayOfWeek.Sunday) cur = cur.AddDays(1); break;
                            case "Weekly": cur = cur.AddDays(7); break;
                            case "Monthly": cur = cur.AddMonths(1); break;
                        }
                    }
                });
                if (ok) Close();
            };
            Controls.AddRange(new Control[] { lblS, _start, lblT, _step, apply, cancel });
            WireButtons(apply, cancel);
        }
    }
}
