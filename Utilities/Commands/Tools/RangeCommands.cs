using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Delete every fully-blank row inside the selection.</summary>
    [ExcelCommand]
    public sealed class DeleteBlankRowsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.DeleteBlankRows",
            Label = "Delete Blank Rows",
            Screentip = "Delete Blank Rows",
            Supertip = "Remove every row in the selection that is completely empty, shifting the remaining rows up.",
            ImageId = "DeleteBlankRows",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range area = ctx.Target.Areas.Count > 1 ? ctx.Target.Areas[1] : ctx.Target;
            Excel.Worksheet ws = area.Worksheet;
            int firstRow = area.Row;
            int rowCount = area.Rows.Count;

            var blankRows = new List<int>();
            for (int i = 0; i < rowCount; i++)
            {
                int sheetRow = firstRow + i;
                Excel.Range row = (Excel.Range)ws.Rows[sheetRow];
                if ((double)ctx.App.WorksheetFunction.CountA(row) == 0)
                    blankRows.Add(sheetRow);
                if ((i & 0x3FF) == 0) ctx.Progress.Report((double)i / rowCount);
            }

            if (blankRows.Count == 0)
            {
                MessageBox.Show("No blank rows found in the selection.", Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            for (int i = blankRows.Count - 1; i >= 0; i--)
                ((Excel.Range)ws.Rows[blankRows[i]]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

            MessageBox.Show(blankRows.Count + " blank row(s) deleted.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Insert a blank row after each row of the selection (alternating rows).</summary>
    [ExcelCommand]
    public sealed class InsertBlankRowsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.InsertBlankRows",
            Label = "Insert Blank Rows",
            Screentip = "Insert Blank Rows",
            Supertip = "Insert one empty row beneath each row in the selection — handy for spacing out a list before printing.",
            ImageId = "InsertBlankRows",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 31,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range area = ctx.Target.Areas.Count > 1 ? ctx.Target.Areas[1] : ctx.Target;
            Excel.Worksheet ws = area.Worksheet;
            int firstRow = area.Row;
            int rowCount = area.Rows.Count;

            // Insert from the bottom up so row indices stay valid; skip the last row's gap.
            for (int i = rowCount - 1; i >= 1; i--)
            {
                Excel.Range insertAt = (Excel.Range)ws.Rows[firstRow + i];
                insertAt.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
            }

            MessageBox.Show((rowCount - 1) + " blank row(s) inserted.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>Select the empty cells within the selection.</summary>
    [ExcelCommand]
    public sealed class SelectBlankCellsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.BlankCells",
            Label = "Select Blank Cells",
            Screentip = "Select Blank Cells",
            Supertip = "Select all empty cells inside the current selection so you can fill or format them at once.",
            ImageId = "SelectBlankCells",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            try
            {
                Excel.Range blanks = ctx.Target.SpecialCells(Excel.XlCellType.xlCellTypeBlanks);
                blanks.Select();
            }
            catch
            {
                MessageBox.Show("No blank cells found in the selection.", Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    /// <summary>Select the non-empty cells (constants and formulas) within the selection.</summary>
    [ExcelCommand]
    public sealed class SelectNonBlankCellsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.NonBlankCells",
            Label = "Select Non-Blank Cells",
            Screentip = "Select Non-Blank Cells",
            Supertip = "Select every cell in the selection that contains a value or formula.",
            ImageId = "SelectNonBlankCells",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 21,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range result = null;
            result = TryUnion(ctx, result, Excel.XlCellType.xlCellTypeConstants);
            result = TryUnion(ctx, result, Excel.XlCellType.xlCellTypeFormulas);

            if (result != null) result.Select();
            else MessageBox.Show("No non-blank cells found.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static Excel.Range TryUnion(CommandContext ctx, Excel.Range acc, Excel.XlCellType type)
        {
            try
            {
                Excel.Range part = ctx.Target.SpecialCells(type);
                return acc == null ? part : ctx.App.Union(acc, part);
            }
            catch { return acc; }
        }
    }

    /// <summary>Select cells that contain a formula error within the selection.</summary>
    [ExcelCommand]
    public sealed class SelectErrorCellsCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.ErrorCells",
            Label = "Select Error Cells",
            Screentip = "Select Error Cells",
            Supertip = "Select all cells in the selection whose formulas return an error (#N/A, #VALUE!, #REF! …).",
            ImageId = "SelectErrorCells",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 22,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            try
            {
                Excel.Range errors = ctx.Target.SpecialCells(
                    Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);
                errors.Select();
            }
            catch
            {
                MessageBox.Show("No error cells found in the selection.", Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
