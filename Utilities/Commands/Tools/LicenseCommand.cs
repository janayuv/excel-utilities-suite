using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using utilities.Dialogs;
using utilities.Services;

namespace utilities.Commands.Tools
{
    [ExcelCommand]
    public sealed class ActivateLicenseCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Suite.ActivateLicense", Label = "Activate License",
            Screentip = "Activate License",
            Supertip = "Enter a product key to unlock the full Excel Utilities Suite.",
            ImageMso = "Permissions",
            Tab = null, Group = "License", Order = 10,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new LicenseDialog(); }
    }

    [ExcelCommand]
    public sealed class DeactivateLicenseCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Suite.DeactivateLicense", Label = "Deactivate",
            Screentip = "Deactivate License",
            Supertip = "Remove the stored product key from this computer.",
            ImageMso = "Delete",
            Tab = null, Group = "License", Order = 20,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new DeactivateDialog(); }
    }

    // ── Activation dialog ─────────────────────────────────────────────────────

    internal sealed class LicenseDialog : DialogBase
    {
        private readonly Label   _statusLbl  = new Label { AutoSize = true };
        // ReadOnly TextBox so user can Ctrl+A / Ctrl+C — prevents transcription errors
        private readonly TextBox _machineLbl = new TextBox { ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = System.Drawing.SystemColors.Control, ForeColor = Color.Gray, Font = new System.Drawing.Font("Courier New", 9f) };
        private readonly TextBox _keyBox     = new TextBox { Width = 230, CharacterCasing = CharacterCasing.Upper };
        private readonly Label   _errorLbl   = new Label { AutoSize = true, ForeColor = Color.Red, Visible = false };
        private bool _formatting;

        public LicenseDialog()
        {
            Text       = "Activate Excel Utilities Suite";
            ClientSize = new Size(380, 212);

            var lblS = new Label { Text = "Current status:", Left = 12, Top = 14, AutoSize = true };
            _statusLbl.SetBounds(120, 14, 240, 16);
            _statusLbl.Text = License.Current.StatusText;

            var lblM   = new Label { Text = "Machine ID:", Left = 12, Top = 34, AutoSize = true };
            _machineLbl.SetBounds(120, 33, 190, 18);
            _machineLbl.Text = RealLicenseService.MachineId();
            _machineLbl.GotFocus += (s, e) => _machineLbl.SelectAll();

            var btnCopy = new Button { Text = "📋 Copy", Left = 316, Top = 30, Width = 52, Height = 22 };
            btnCopy.Click += (s, e) =>
            {
                Clipboard.SetText(_machineLbl.Text);
                btnCopy.Text = "✓ Copied";
                var t = new System.Windows.Forms.Timer { Interval = 1500 };
                t.Tick += (ts, te) => { btnCopy.Text = "📋 Copy"; t.Stop(); t.Dispose(); };
                t.Start();
            };

            var sep  = new Label { BorderStyle = BorderStyle.Fixed3D, Left = 12, Top = 58, Width = 356, Height = 2 };
            var lblK = new Label { Text = "Product key:", Left = 12, Top = 72, AutoSize = true };
            _keyBox.SetBounds(12, 92, 356, 23);
            _keyBox.TextChanged += (s, e) => { _errorLbl.Visible = false; FormatKey(); };

            _errorLbl.SetBounds(12, 120, 356, 16);

            var activate = new Button { Text = "&Activate", Left = 196, Top = 170, Width = 80 };
            var cancel   = new Button { Text = "&Cancel",   Left = 288, Top = 170, Width = 80, DialogResult = DialogResult.Cancel };
            activate.Click += OnActivate;

            Controls.AddRange(new Control[] { lblS, _statusLbl, lblM, _machineLbl, btnCopy, sep, lblK, _keyBox, _errorLbl, activate, cancel });
            WireButtons(activate, cancel);
            ActiveControl = _keyBox;
        }

        private void FormatKey()
        {
            if (_formatting) return;
            _formatting = true;
            try
            {
                string raw = _keyBox.Text.Replace("-", "");
                if (raw.Length > 20) raw = raw.Substring(0, 20);
                var sb = new StringBuilder();
                for (int i = 0; i < raw.Length; i++)
                {
                    if (i > 0 && i % 5 == 0) sb.Append('-');
                    sb.Append(raw[i]);
                }
                string fmt = sb.ToString();
                if (fmt != _keyBox.Text)
                {
                    int caret = _keyBox.SelectionStart;
                    _keyBox.Text = fmt;
                    _keyBox.SelectionStart = Math.Min(caret, fmt.Length);
                }
            }
            finally { _formatting = false; }
        }

        private void OnActivate(object sender, EventArgs e)
        {
            string key = _keyBox.Text.Trim();
            if (key.Length == 0) { ShowErr("Enter a product key."); return; }

            var real = License.Current as RealLicenseService;
            if (real == null)
            {
                if (!RealLicenseService.ValidateKey(key))
                { ShowErr("Invalid product key. Please check and try again."); return; }
                MessageBox.Show("Key format valid (running in stub mode — not persisted).",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close(); return;
            }

            if (!real.Activate(key))
            { ShowErr("Invalid product key. Please check and try again."); return; }

            MessageBox.Show("Activation successful! Thank you for your purchase.",
                Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        private void ShowErr(string msg) { _errorLbl.Text = msg; _errorLbl.Visible = true; }
    }

    // ── Deactivation dialog ───────────────────────────────────────────────────

    internal sealed class DeactivateDialog : DialogBase
    {
        public DeactivateDialog()
        {
            Text       = "Deactivate License";
            ClientSize = new Size(360, 134);

            var lbl = new Label
            {
                Text = "Remove the product key from this computer?\n\n" +
                       "You can re-activate later using the same key.\n" +
                       "Current: " + License.Current.StatusText,
                Left = 12, Top = 12, Width = 336, Height = 64, AutoSize = false
            };
            var remove = new Button { Text = "&Remove Key", Left = 168, Top = 90, Width = 90 };
            var cancel = new Button { Text = "&Cancel",     Left = 268, Top = 90, Width = 80, DialogResult = DialogResult.Cancel };
            remove.Click += (s, e) =>
            {
                (License.Current as RealLicenseService)?.Deactivate();
                MessageBox.Show("License removed from this computer.", Text,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                Close();
            };
            Controls.AddRange(new Control[] { lbl, remove, cancel });
            WireButtons(remove, cancel);
        }
    }
}
