using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public static class ExcelHelper
    {
        public static void ColorDuplicatesOnlyForDuplicates()
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a range of cells first.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var colorList = new List<object>
            {
                Excel.XlRgbColor.rgbSkyBlue,
                Excel.XlRgbColor.rgbLightGreen,
                Excel.XlRgbColor.rgbLightPink,
                Excel.XlRgbColor.rgbLightYellow,
                Excel.XlRgbColor.rgbLightCoral,
                Excel.XlRgbColor.rgbLavender,
                Excel.XlRgbColor.rgbPaleGreen,
                Excel.XlRgbColor.rgbWheat,
                Excel.XlRgbColor.rgbLightSalmon,
                Excel.XlRgbColor.rgbKhaki
            };

            var colorDict = new Dictionary<string, object>();
            var countDict = new Dictionary<string, int>();

            foreach (Excel.Range cell in selectedRange)
            {
                if (cell.Value2 != null && !string.IsNullOrEmpty(cell.Value2.ToString()))
                {
                    string value = cell.Value2.ToString();
                    if (countDict.ContainsKey(value))
                        countDict[value]++;
                    else
                        countDict.Add(value, 1);
                }
            }

            int colorIndex = 0;
            foreach (Excel.Range cell in selectedRange)
            {
                if (cell.Value2 != null && !string.IsNullOrEmpty(cell.Value2.ToString()))
                {
                    string value = cell.Value2.ToString();
                    if (countDict[value] > 1)
                    {
                        if (!colorDict.ContainsKey(value))
                        {
                            if (colorIndex < colorList.Count)
                            {
                                colorDict[value] = colorList[colorIndex++];
                            }
                            else
                            {
                                colorDict[value] = GenerateUniqueColor(colorIndex++); // Generate dynamic color
                            }
                        }

                        cell.Interior.Color = colorDict[value];
                    }
                    else
                    {
                        cell.Interior.ColorIndex = Excel.Constants.xlNone;
                    }
                }
            }

            MessageBox.Show("Duplicates have been successfully colored.", "Operation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static object GenerateUniqueColor(int index)
        {
            Random rnd = new Random(index);
            int red = rnd.Next(50, 255);   // Avoid very dark colors
            int green = rnd.Next(50, 255);
            int blue = rnd.Next(50, 255);

            return (red << 16) | (green << 8) | blue; // Combine RGB values
        }
    }
}
