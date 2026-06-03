using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>
    /// Convert text that looks like numbers into real numeric values. Port of the original
    /// DataConverterHelper.ConvertTextToNumbers, now guarded/undoable via the framework.
    /// </summary>
    [ExcelCommand]
    public sealed class ConvertTextToNumbersCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.TextToNumbers",
            Label = "Text to Numbers",
            Screentip = "Convert Text to Numbers",
            Supertip = "Turn numbers stored as text into real numeric values so they can be summed, sorted and used in formulas.",
            ImageId = "ConvertTextToNumbers",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int total = ctx.Target.Cells.Count;
            int done = 0;
            int converted = 0;

            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v != null)
                {
                    double number;
                    if (double.TryParse(v.ToString(), out number))
                    {
                        cell.Value2 = number;
                        converted++;
                    }
                }
                done++;
                if ((done & 0x3FF) == 0) ctx.Progress.Report((double)done / total);
            }

            ctx.Progress.Report(1);
            MessageBox.Show(converted + " value(s) converted to numbers.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    /// <summary>
    /// Convert numeric values into text (formatted as text, leading apostrophe). Port of the
    /// original DataConverterHelper.ConvertNumbersToText.
    /// </summary>
    [ExcelCommand]
    public sealed class ConvertNumbersToTextCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Convert.NumbersToText",
            Label = "Numbers to Text",
            Screentip = "Convert Numbers to Text",
            Supertip = "Store numeric values as text (with a Text cell format) to preserve leading zeros and long IDs.",
            ImageId = "ConvertNumbersToText",
            Tab = "Editing",
            Group = "Format & Convert",
            Order = 11,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        protected override void Run(CommandContext ctx)
        {
            int total = ctx.Target.Cells.Count;
            int done = 0;
            int converted = 0;

            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v != null && v is double)
                {
                    cell.NumberFormat = "@";
                    cell.Value2 = "'" + v.ToString();
                    converted++;
                }
                done++;
                if ((done & 0x3FF) == 0) ctx.Progress.Report((double)done / total);
            }

            ctx.Progress.Report(1);
            MessageBox.Show(converted + " value(s) converted to text.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
