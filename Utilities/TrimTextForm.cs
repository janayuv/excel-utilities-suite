using System;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public partial class TrimTextForm : Form
    {
        public TrimTextForm()
        {
            InitializeComponent();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            var selectedRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            if (selectedRange == null)
            {
                MessageBox.Show("Please select a valid range.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Handle the selected option
            if (optDeleteLeadingTrailingSpaces.Checked)
            {
                DeleteLeadingTrailingSpaces(selectedRange);
            }
            else if (optDeleteLeadingTrailingExcessiveSpaces.Checked)
            {
                DeleteLeadingTrailingExcessiveSpaces(selectedRange);
            }
            else if (optDeleteLeadingCharacters.Checked)
            {
                DeleteLeadingCharacters(selectedRange);
            }
            else if (optDeleteEndingCharacters.Checked)
            {
                DeleteEndingCharacters(selectedRange);
            }
            else
            {
                MessageBox.Show("Please select an option.", "No Option Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            MessageBox.Show("Operation completed successfully.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Delete leading and trailing spaces
        private void DeleteLeadingTrailingSpaces(Excel.Range selectedRange)
        {
            foreach (Excel.Range cell in selectedRange.Cells)
            {
                if (cell.Value2 != null && cell.Value2 is string)
                {
                    cell.Value2 = cell.Value2.ToString().Trim();
                }
            }
        }

        // Delete leading, trailing, and excessive spaces
        private void DeleteLeadingTrailingExcessiveSpaces(Excel.Range selectedRange)
        {
            foreach (Excel.Range cell in selectedRange.Cells)
            {
                if (cell.Value2 != null && cell.Value2 is string)
                {
                    string cleanedText = System.Text.RegularExpressions.Regex.Replace(cell.Value2.ToString().Trim(), @"\s+", " ");
                    cell.Value2 = cleanedText;
                }
            }
        }

        // Delete a specified number of leading characters
        private void DeleteLeadingCharacters(Excel.Range selectedRange)
        {
            // Validate input from the TextBox
            int numberOfChars;
            if (int.TryParse(txtNumberOfCharacters.Text, out numberOfChars) && numberOfChars > 0)
            {
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    if (cell.Value2 != null && cell.Value2 is string)
                    {
                        string newValue = cell.Value2.ToString();
                        if (newValue.Length > numberOfChars)
                        {
                            newValue = newValue.Substring(numberOfChars);
                        }
                        cell.Value2 = newValue;
                    }
                }
            }
            else
            {
                MessageBox.Show("Please enter a valid number of characters to remove.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // Delete a specified number of ending characters
        private void DeleteEndingCharacters(Excel.Range selectedRange)
        {
            // Validate input from the TextBox
            int numberOfChars;
            if (int.TryParse(txtNumberOfCharacters.Text, out numberOfChars) && numberOfChars > 0)
            {
                foreach (Excel.Range cell in selectedRange.Cells)
                {
                    if (cell.Value2 != null && cell.Value2 is string)
                    {
                        string newValue = cell.Value2.ToString();
                        if (newValue.Length > numberOfChars)
                        {
                            newValue = newValue.Substring(0, newValue.Length - numberOfChars);
                        }
                        cell.Value2 = newValue;
                    }
                }
            }
            else
            {
                MessageBox.Show("Please enter a valid number of characters to remove.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string TruncateString(string input, int numberOfChars, bool isLeading)
        {
            if (isLeading)
                return input.Length > numberOfChars ? input.Substring(numberOfChars) : input;
            else
                return input.Length > numberOfChars ? input.Substring(0, input.Length - numberOfChars) : input;
        }

        // Helper method to get number of characters from user
        private int GetNumberInput(string prompt)
        {
            using (var inputBox = new Form())
            {
                var label = new Label() { Left = 20, Top = 20, Text = prompt };
                var textBox = new TextBox() { Left = 20, Top = 50, Width = 200 };
                var okButton = new Button() { Text = "OK", Left = 20, Width = 75, Top = 80 };
                var cancelButton = new Button() { Text = "Cancel", Left = 120, Width = 75, Top = 80 };

                inputBox.Controls.Add(label);
                inputBox.Controls.Add(textBox);
                inputBox.Controls.Add(okButton);
                inputBox.Controls.Add(cancelButton);

                okButton.DialogResult = DialogResult.OK;
                cancelButton.DialogResult = DialogResult.Cancel;
                inputBox.StartPosition = FormStartPosition.CenterScreen;

                inputBox.AcceptButton = okButton;
                inputBox.CancelButton = cancelButton;

                if (inputBox.ShowDialog() == DialogResult.OK)
                {
                    int number;
                    if (int.TryParse(textBox.Text, out number))
                    {
                        return number;
                    }
                    else
                    {
                        MessageBox.Show("Invalid number entered. Please try again.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return 0;
                    }
                }
                else
                {
                    return 0;
                }
            }
        }
    }
}
