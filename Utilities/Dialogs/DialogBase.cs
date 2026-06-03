using System;
using System.Drawing;
using System.Windows.Forms;
using utilities.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Dialogs
{
    /// <summary>
    /// Common base for every tool dialog. Enforces the suite's usability standard
    /// (consistent chrome, Enter/Esc handling, DPI scaling, inline validation) and routes
    /// the actual mutation through <see cref="OperationRunner"/> so undo and error handling
    /// are identical to direct commands.
    /// </summary>
    public class DialogBase : Form
    {
        /// <summary>Inline (non-blocking) validation indicator placed next to offending fields.</summary>
        protected readonly ErrorProvider Validation;

        public DialogBase()
        {
            Validation = new ErrorProvider { BlinkStyle = ErrorBlinkStyle.NeverBlink };

            // Consistent chrome and DPI behaviour. Designer-set values in derived forms win
            // where they differ, but these give every dialog a uniform baseline.
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            KeyPreview = true; // so Esc closes even when a child control has focus
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                e.Handled = true;
                return;
            }
            base.OnKeyDown(e);
        }

        /// <summary>
        /// Wire the standard primary/secondary buttons: Enter triggers <paramref name="accept"/>,
        /// Esc triggers <paramref name="cancel"/>. Call from the derived form's load/ctor.
        /// </summary>
        protected void WireButtons(IButtonControl accept, IButtonControl cancel)
        {
            if (accept != null) AcceptButton = accept;
            if (cancel != null) CancelButton = cancel;
        }

        /// <summary>Show or clear an inline validation message for a control.</summary>
        protected void SetError(Control control, string message)
        {
            Validation.SetError(control, message ?? string.Empty);
        }

        /// <summary>True when no control currently has a validation error set.</summary>
        protected bool HasNoErrors(params Control[] controls)
        {
            foreach (Control c in controls)
            {
                if (!string.IsNullOrEmpty(Validation.GetError(c))) return false;
            }
            return true;
        }

        /// <summary>The currently selected range, or null.</summary>
        protected Excel.Range CurrentSelection
        {
            get { return Globals.ThisAddIn.Application.Selection as Excel.Range; }
        }

        /// <summary>
        /// Run a guarded operation for this dialog using its command definition. Returns true
        /// on success. Shows the standard "select a range" warning when a target is required
        /// but missing.
        /// </summary>
        protected bool RunOperation(CommandDefinition def, Excel.Range target, Action<CommandContext> work)
        {
            if (def.RequiresSelection && target == null)
            {
                MessageBox.Show("Please select a range of cells first.", def.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return OperationRunner.RunGuarded(def, target, work);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && Validation != null) Validation.Dispose();
            base.Dispose(disposing);
        }
    }
}
