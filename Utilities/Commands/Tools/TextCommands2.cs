using System;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Reverse the character order of the text in every selected cell.</summary>
    [ExcelCommand]
    public sealed class ReverseTextCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.Reverse",
            Label = "Reverse Text",
            Screentip = "Reverse Text",
            Supertip = "Reverse the character order of the text in each selected cell (e.g. \"Excel\" becomes \"lecxE\").",
            ImageId = "ReverseText",
            Tab = "Editing",
            Group = "Text",
            Order = 30,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                char[] chars = v.ToString().ToCharArray();
                Array.Reverse(chars);
                cell.Value2 = new string(chars);
            }
        }
    }

    /// <summary>Keep only the digits in each selected cell (extract numbers).</summary>
    [ExcelCommand]
    public sealed class ExtractNumbersCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.ExtractNumbers",
            Label = "Extract Numbers",
            Screentip = "Extract Numbers",
            Supertip = "Strip everything except the digits from each cell, leaving only the numeric characters.",
            ImageId = "ExtractNumbers",
            Tab = "Editing",
            Group = "Text",
            Order = 31,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                var sb = new StringBuilder();
                foreach (char c in v.ToString())
                    if (char.IsDigit(c)) sb.Append(c);
                cell.Value2 = sb.ToString();
            }
        }
    }

    /// <summary>Keep only the letters in each selected cell (extract text).</summary>
    [ExcelCommand]
    public sealed class ExtractTextCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.ExtractText",
            Label = "Extract Text",
            Screentip = "Extract Text",
            Supertip = "Strip everything except the letters from each cell, leaving only the alphabetic characters.",
            ImageId = "ExtractText",
            Tab = "Editing",
            Group = "Text",
            Order = 32,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                var sb = new StringBuilder();
                foreach (char c in v.ToString())
                    if (char.IsLetter(c)) sb.Append(c);
                cell.Value2 = sb.ToString();
            }
        }
    }

    /// <summary>
    /// Split full names in a single-column selection into first and last name columns,
    /// inserting two columns to the right.
    /// </summary>
    [ExcelCommand]
    public sealed class SplitNamesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.SplitNames",
            Label = "Split Names",
            Screentip = "Split Names",
            Supertip = "Split full names in a single column into separate First name and Last name columns inserted to the right.",
            ImageId = "SplitNames",
            Tab = "Editing",
            Group = "Text",
            Order = 33,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range area = ctx.Target.Areas.Count > 1 ? ctx.Target.Areas[1] : ctx.Target;
            if (area.Columns.Count != 1)
            {
                MessageBox.Show("Select a single column of full names.", Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Excel.Worksheet ws = area.Worksheet;
            int col = area.Column;

            // Insert two columns to the right for First / Last.
            ((Excel.Range)ws.Columns[col + 1]).Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            ((Excel.Range)ws.Columns[col + 1]).Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);

            int firstRow = area.Row;
            int count = area.Rows.Count;
            for (int i = 0; i < count; i++)
            {
                int r = firstRow + i;
                object v = ((Excel.Range)ws.Cells[r, col]).Value2;
                string full = v != null ? v.ToString().Trim() : string.Empty;
                string first = full, last = string.Empty;
                int sp = full.IndexOf(' ');
                if (sp >= 0)
                {
                    first = full.Substring(0, sp);
                    last = full.Substring(sp + 1).Trim();
                }
                ((Excel.Range)ws.Cells[r, col + 1]).Value2 = first;
                ((Excel.Range)ws.Cells[r, col + 2]).Value2 = last;
            }

            MessageBox.Show("Names split into First and Last columns.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
