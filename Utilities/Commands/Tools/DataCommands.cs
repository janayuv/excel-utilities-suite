using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>
    /// Delete rows in the selection whose full contents duplicate an earlier row, keeping the
    /// first occurrence. Destructive, so it confirms first; deletion is not value-undoable.
    /// </summary>
    [ExcelCommand]
    public sealed class RemoveDuplicateRowsCommand : CommandBase
    {
        private const char FieldSeparator = ''; // unit separator, unlikely to appear in cells

        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Data.RemoveDuplicateRows",
            Label = "Remove Duplicate Rows",
            Screentip = "Remove Duplicate Rows",
            Supertip = "Delete rows in the selection that exactly repeat an earlier row, keeping the first occurrence of each.",
            ImageId = "RemoveDuplicateRows",
            Tab = "Data & Cleaning",
            Group = "Duplicates",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range target = ctx.Target;
            Excel.Range area = target.Areas.Count > 1 ? target.Areas[1] : target;

            object[,] values = area.Value2 as object[,];
            if (values == null)
            {
                MessageBox.Show("Select at least two rows.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int rows = values.GetLength(0);
            int cols = values.GetLength(1);

            var seen = new HashSet<string>(StringComparer.Ordinal);
            var duplicateRowOffsets = new List<int>(); // 1-based offsets within the area

            for (int r = 1; r <= rows; r++)
            {
                var sb = new StringBuilder();
                for (int c = 1; c <= cols; c++)
                {
                    object v = values[r, c];
                    sb.Append(v == null ? " " : v.ToString());
                    sb.Append(FieldSeparator);
                }
                string sig = sb.ToString();
                if (!seen.Add(sig)) duplicateRowOffsets.Add(r);
            }

            if (duplicateRowOffsets.Count == 0)
            {
                MessageBox.Show("No duplicate rows found.", Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult confirm = MessageBox.Show(
                duplicateRowOffsets.Count + " duplicate row(s) will be deleted. This cannot be undone with the Undo button. Continue?",
                Definition.Label, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            // Delete from the bottom up so earlier offsets stay valid.
            int firstRow = area.Row;
            Excel.Worksheet ws = area.Worksheet;
            for (int i = duplicateRowOffsets.Count - 1; i >= 0; i--)
            {
                int sheetRow = firstRow + duplicateRowOffsets[i] - 1;
                ((Excel.Range)ws.Rows[sheetRow]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }

            MessageBox.Show(duplicateRowOffsets.Count + " duplicate row(s) removed.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
