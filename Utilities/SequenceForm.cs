using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace utilities
{
    public partial class SequenceForm : Form
    {
        public enum FillOrder
        {
            Vertical,
            Horizontal
        }

        public int StartNumber { get; private set; }
        public int Increment { get; private set; }
        public int NumberOfDigits { get; private set; }
        public string Prefix { get; private set; }
        public string Suffix { get; private set; }
        public FillOrder SequenceFillOrder { get; private set; }

        public SequenceForm()
        {
            InitializeComponent();

            // Wire up the event handlers
            btnGeneratePreview.Click += btnGeneratePreview_Click;
            btnOK.Click += btnOK_Click;
            btnCancel.Click += btnCancel_Click;

            // Validate inputs dynamically
            txtStartNumber.TextChanged += (s, e) => ValidateInputs();
            txtIncrement.TextChanged += (s, e) => ValidateInputs();
            comboBoxFillOrder.SelectedIndexChanged += (s, e) => ValidateInputs();
        }

        private void SequenceForm_Load(object sender, EventArgs e)
        {
            // Set default fill order to "Fill Horizontally"
            if (comboBoxFillOrder.Items.Count == 0) // Ensure items are not already added
            {
                comboBoxFillOrder.Items.Add("Fill Vertically");
                comboBoxFillOrder.Items.Add("Fill Horizontally");
            }
            comboBoxFillOrder.SelectedItem = "Fill Horizontally"; // Default value

            // Set default number of digits to 0
            if (string.IsNullOrWhiteSpace(txtDigits.Text))
                txtDigits.Text = "0";

            ValidateInputs(); // Ensure inputs are valid on load
        }

        private void btnGeneratePreview_Click(object sender, EventArgs e)
        {
            try
            {
                // Use local variables to store parsed values
                int startNumber, increment, numberOfDigits;

                if (!int.TryParse(txtStartNumber.Text, out startNumber))
                    throw new Exception("Invalid start number.");
                if (!int.TryParse(txtIncrement.Text, out increment))
                    throw new Exception("Invalid increment value.");
                if (!int.TryParse(txtDigits.Text, out numberOfDigits))
                    throw new Exception("Invalid number of digits.");

                // Assign to properties after successful parsing
                StartNumber = startNumber;
                Increment = increment;
                NumberOfDigits = numberOfDigits;

                Prefix = txtPrefix.Text;
                Suffix = txtSuffix.Text;

                if (comboBoxFillOrder.SelectedItem == null)
                    throw new Exception("Fill order is not selected.");

                SequenceFillOrder = comboBoxFillOrder.SelectedItem.ToString() == "Fill Vertically"
                    ? FillOrder.Vertical
                    : FillOrder.Horizontal;

                // Generate preview
                listBoxPreview.Items.Clear();
                int numberOfEntries = 10; // Default number of entries in preview
                for (int i = 0; i < numberOfEntries; i++)
                {
                    int currentNumber = StartNumber + (i * Increment);
                    string formattedNumber = Prefix + currentNumber.ToString().PadLeft(NumberOfDigits > 0 ? NumberOfDigits : currentNumber.ToString().Length, '0') + Suffix;
                    listBoxPreview.Items.Add(formattedNumber);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Input Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                // Get the Excel application and active worksheet
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;

                if (activeSheet == null)
                {
                    MessageBox.Show("No active sheet found.");
                    return;
                }

                // Get the selected range of cells in the active sheet
                Excel.Range selectedRange = excelApp.Selection as Excel.Range;

                if (selectedRange == null || selectedRange.Cells.Count == 0)
                {
                    MessageBox.Show("Please select some cells to insert the sequence.");
                    return;
                }

                // Get only visible cells in the selected range
                Excel.Range visibleCells = selectedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                // Initialize the sequence variables
                int currentNumber = StartNumber;

                // Loop through visible cells and apply the sequence
                foreach (Excel.Range cell in visibleCells)
                {
                    string formattedNumber = Prefix + currentNumber.ToString().PadLeft(NumberOfDigits > 0 ? NumberOfDigits : currentNumber.ToString().Length, '0') + Suffix;
                    cell.Value = formattedNumber;
                    currentNumber += Increment;
                }

                // Show a confirmation message
                MessageBox.Show("Sequence inserted into visible cells!");

                // Close the form
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting sequence: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            // Handle Cancel button click
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void ValidateInputs()
        {
            // Enable OK button only if all inputs are valid
            btnOK.Enabled = !string.IsNullOrWhiteSpace(txtStartNumber.Text) &&
                            !string.IsNullOrWhiteSpace(txtIncrement.Text) &&
                            comboBoxFillOrder.SelectedItem != null;
        }
    }
}