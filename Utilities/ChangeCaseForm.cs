using System;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public partial class ChangeCaseForm : Form
    {
        public ChangeCaseForm()
        {
            InitializeComponent();
        }

        // Apply the selected case transformation
        private void btnApply_Click(object sender, EventArgs e)
        {
            var selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (Excel.Range cell in selectedRange.Cells)
            {
                if (cell.Value2 != null && cell.Value2 is string)
                {
                    string newValue = cell.Value2.ToString();
                    if (optLowercase.Checked) newValue = newValue.ToLower();
                    else if (optUppercase.Checked) newValue = newValue.ToUpper();
                    else if (optToggleCase.Checked) newValue = ToggleCase(newValue);
                    else if (optSentenceCase.Checked) newValue = SentenceCase(newValue);
                    else if (optGlossaryCase.Checked) newValue = GlossaryCase(newValue);
                    else if (optCapitalizeEachWord.Checked) newValue = CapitalizeEachWord(newValue);
                    else if (optStartEachWordWithUppercase.Checked) newValue = CapitalizeEachWord(newValue);

                    cell.Value2 = newValue;
                }
            }

            MessageBox.Show("Case transformation applied successfully!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        // Close the form when cancel is clicked
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Helper methods for case transformations
        private string ToggleCase(string text)
        {
            return new string(text.Select(c => char.IsLower(c) ? char.ToUpper(c) : char.ToLower(c)).ToArray());
        }

        private string SentenceCase(string text)
        {
            return string.Join(". ", text.Split(new[] { ". " }, StringSplitOptions.None)
                .Select(sentence => char.ToUpper(sentence[0]) + sentence.Substring(1).ToLower()));
        }

        private string GlossaryCase(string text)
        {
            return string.Join(". ", text.Split(new[] { ". " }, StringSplitOptions.None)
                .Select(sentence => char.ToLower(sentence[0]) + sentence.Substring(1)));
        }

        private string CapitalizeEachWord(string text)
        {
            return string.Join(" ", text.Split(' ')
                .Select(word => char.ToUpper(word[0]) + word.Substring(1).ToLower()));
        }
    }
}
