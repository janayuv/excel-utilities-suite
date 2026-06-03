using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using Microsoft.Office.Core;
using System;

namespace utilities
{
    public partial class UtilitiesRibbon : RibbonBase
    {
        private void UtilitiesRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Optional: Any load logic for the ribbon
        }
        private void btnColorDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            ExcelHelper.ColorDuplicatesOnlyForDuplicates();
        }


        private void btnInsertSequence_Click(object sender, RibbonControlEventArgs e)
        {
            SequenceForm sequenceForm = new SequenceForm(); // Your form name
            sequenceForm.Show(); // This will display the form
        }
        private void btnConvertTextToNumbers_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange != null)
            {
                DataConverterHelper.ConvertTextToNumbers(selectedRange);
            }
            else
            {
                MessageBox.Show("Please select a valid range.", "Invalid Range", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnConvertNumbersToText_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange != null)
            {
                DataConverterHelper.ConvertNumbersToText(selectedRange);
            }
            else
            {
                MessageBox.Show("Please select a valid range.", "Invalid Range", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnChangeCase_Click(object sender, RibbonControlEventArgs e)
        {
            ChangeCaseForm changeCaseForm = new ChangeCaseForm();
            changeCaseForm.Show();
        }

        // Event handler for btnTrimText
        private void btnTrimText_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                TrimTextForm trimTextForm = new TrimTextForm();
                trimTextForm.ShowDialog(); // This opens the form as a modal dialog
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
