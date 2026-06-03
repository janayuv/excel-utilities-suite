using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public static class DataConverterHelper
    {
        // Convert Text to Numbers
        public static void ConvertTextToNumbers(Excel.Range selectedRange)
        {
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Turn off settings for speed
                Excel.Application app = Globals.ThisAddIn.Application;
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Process cells
                int cellCount = selectedRange.Cells.Count;
                int processedCount = 0;

                foreach (Excel.Range cell in selectedRange)
                {
                    if (cell.Value2 != null)
                    {
                        double number;
                        // Attempt to parse the cell value
                        if (double.TryParse(cell.Value2.ToString(), out number))
                        {
                            // If parsing is successful, assign the number to the cell
                            cell.Value2 = number; // Convert to number
                        }
                        else
                        {
                            // If not a number, show message
                            MessageBox.Show($"Could not convert value '{cell.Value2}' to number.", "Conversion Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }

                    // Update progress
                    processedCount++;
                    app.StatusBar = $"Processing: {Math.Round((double)processedCount / cellCount * 100)}%";
                }

                MessageBox.Show("Text to Numbers conversion complete!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Reset settings
                Excel.Application app = Globals.ThisAddIn.Application;
                app.ScreenUpdating = true;
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.EnableEvents = true;
                app.StatusBar = false;
            }
        }

        // Convert Numbers to Text
        public static void ConvertNumbersToText(Excel.Range selectedRange)
        {
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range.", "Invalid Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                // Turn off settings for speed
                Excel.Application app = Globals.ThisAddIn.Application;
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Process cells
                int cellCount = selectedRange.Cells.Count;
                int processedCount = 0;

                foreach (Excel.Range cell in selectedRange)
                {
                    if (cell.Value2 != null && cell.Value2 is double)
                    {
                        // Convert Numbers to Text by adding a leading apostrophe ('), which Excel treats as text
                        cell.Value2 = "'" + cell.Value2.ToString();

                        // Optional: set the cell format to text explicitly to ensure it displays correctly
                        cell.NumberFormat = "@"; // Set format to Text
                    }

                    // Update progress
                    processedCount++;
                    app.StatusBar = $"Processing: {Math.Round((double)processedCount / cellCount * 100)}%";
                }

                MessageBox.Show("Numbers to Text conversion complete!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Reset settings
                Excel.Application app = Globals.ThisAddIn.Application;
                app.ScreenUpdating = true;
                app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                app.EnableEvents = true;
                app.StatusBar = false;
            }
        }
    }
}
