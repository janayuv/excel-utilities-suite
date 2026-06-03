using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>
    /// Highlight values that appear more than once in the selection, giving each repeated
    /// value its own colour. Port of the original ExcelHelper.ColorDuplicatesOnlyForDuplicates.
    /// </summary>
    [ExcelCommand]
    public sealed class ColorDuplicatesCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Range.ColorDuplicates",
            Label = "Color Duplicates",
            Screentip = "Color Duplicate Values",
            Supertip = "Find values that occur more than once in the selection and shade each duplicate group with its own colour.",
            ImageId = "ColorDuplicates",
            Tab = "Utilities",
            Group = "Range & Cells",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None // colour fills are not captured by the value-based undo
        };

        private static readonly object[] Palette =
        {
            Excel.XlRgbColor.rgbSkyBlue, Excel.XlRgbColor.rgbLightGreen, Excel.XlRgbColor.rgbLightPink,
            Excel.XlRgbColor.rgbLightYellow, Excel.XlRgbColor.rgbLightCoral, Excel.XlRgbColor.rgbLavender,
            Excel.XlRgbColor.rgbPaleGreen, Excel.XlRgbColor.rgbWheat, Excel.XlRgbColor.rgbLightSalmon,
            Excel.XlRgbColor.rgbKhaki
        };

        protected override void Run(CommandContext ctx)
        {
            var counts = new Dictionary<string, int>();
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                string key = v.ToString();
                if (key.Length == 0) continue;
                counts[key] = counts.ContainsKey(key) ? counts[key] + 1 : 1;
            }

            var assigned = new Dictionary<string, object>();
            int colorIndex = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                object v = cell.Value2;
                if (v == null) continue;
                string key = v.ToString();
                if (key.Length == 0) continue;

                if (counts.ContainsKey(key) && counts[key] > 1)
                {
                    object color;
                    if (!assigned.TryGetValue(key, out color))
                    {
                        color = colorIndex < Palette.Length ? Palette[colorIndex] : RandomColor(colorIndex);
                        assigned[key] = color;
                        colorIndex++;
                    }
                    cell.Interior.Color = color;
                }
                else
                {
                    cell.Interior.ColorIndex = Excel.Constants.xlNone;
                }
            }

            MessageBox.Show(assigned.Count + " duplicate group(s) highlighted.", Definition.Label,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static object RandomColor(int seed)
        {
            var rnd = new Random(seed);
            int r = rnd.Next(50, 255), g = rnd.Next(50, 255), b = rnd.Next(50, 255);
            return (r << 16) | (g << 8) | b;
        }
    }

    /// <summary>Select every cell in the used range whose fill colour matches the active cell.</summary>
    [ExcelCommand]
    public sealed class SelectByColorCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Select.ByColor",
            Label = "Select by Color",
            Screentip = "Select Cells by Fill Color",
            Supertip = "Select all cells on the sheet that share the same background colour as the active cell.",
            ImageId = "SelectByColor",
            Tab = "Select & Navigate",
            Group = "Select",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range active = ctx.App.ActiveCell;
            if (active == null) { MessageBox.Show("Click a cell with the colour to match.", Definition.Label); return; }

            object targetColor = active.Interior.Color;
            Excel.Worksheet ws = ctx.App.ActiveSheet as Excel.Worksheet;
            Excel.Range used = ws != null ? ws.UsedRange : ctx.Target;

            Excel.Range matches = null;
            int total = used.Cells.Count, done = 0;
            foreach (Excel.Range cell in used.Cells)
            {
                if (Equals(cell.Interior.Color, targetColor))
                    matches = matches == null ? cell : ctx.App.Union(matches, cell);
                if ((++done & 0x3FF) == 0) ctx.Progress.Report((double)done / total);
            }

            if (matches != null) matches.Select();
            else MessageBox.Show("No other cells share that colour.", Definition.Label);
        }
    }

    /// <summary>Sum and count the selected cells whose fill colour matches the active cell.</summary>
    [ExcelCommand]
    public sealed class SumCountByColorCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Formula.SumCountByColor",
            Label = "Sum/Count by Color",
            Screentip = "Sum and Count by Color",
            Supertip = "Total and count the selected cells that share the active cell's fill colour, without writing a helper formula.",
            ImageId = "SumCountByColor",
            Tab = "Formula & Statistics",
            Group = "Statistics",
            Order = 10,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Range active = ctx.App.ActiveCell;
            object targetColor = active != null ? active.Interior.Color : null;

            double sum = 0;
            int count = 0;
            foreach (Excel.Range cell in ctx.Target.Cells)
            {
                if (!Equals(cell.Interior.Color, targetColor)) continue;
                count++;
                object v = cell.Value2;
                double d;
                if (v != null && double.TryParse(v.ToString(), NumberStyles.Any, CultureInfo.CurrentCulture, out d))
                    sum += d;
            }

            MessageBox.Show(
                "Cells matching the active colour:\n\nCount: " + count + "\nSum: " + sum.ToString("0.##"),
                Definition.Label, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
