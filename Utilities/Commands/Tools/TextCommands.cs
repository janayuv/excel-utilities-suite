using System;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using utilities.Dialogs;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>Add a prefix, suffix or insert text at a position across the selection.</summary>
    [ExcelCommand]
    public sealed class AddTextCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Text.AddText",
            Label = "Add Text",
            Screentip = "Add Text",
            Supertip = "Insert text as a prefix, a suffix, or at a specific character position in every selected cell.",
            ImageId = "AddText",
            Tab = "Editing",
            Group = "Text",
            Order = 20,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new AddTextDialog(); }
    }

    internal sealed class AddTextDialog : DialogBase
    {
        private readonly TextBox _text = new TextBox();
        private readonly RadioButton _prefix = new RadioButton { Text = "Before first character (prefix)", Checked = true };
        private readonly RadioButton _suffix = new RadioButton { Text = "After last character (suffix)" };
        private readonly RadioButton _at = new RadioButton { Text = "At position:" };
        private readonly NumericUpDown _pos = new NumericUpDown { Minimum = 1, Maximum = 9999, Value = 1, Enabled = false };

        public AddTextDialog()
        {
            Text = AddTextCommand.Def.Label;
            ClientSize = new Size(360, 200);

            var lblText = new Label { Text = "Text:", Left = 12, Top = 15, Width = 60, AutoSize = true };
            _text.SetBounds(80, 12, 260, 23);

            _prefix.SetBounds(80, 48, 260, 22);
            _suffix.SetBounds(80, 72, 260, 22);
            _at.SetBounds(80, 96, 90, 22);
            _pos.SetBounds(176, 96, 60, 23);
            _at.CheckedChanged += (s, e) => _pos.Enabled = _at.Checked;

            var apply = new Button { Text = "&Apply", Left = 176, Top = 150, Width = 80, DialogResult = DialogResult.None };
            var cancel = new Button { Text = "&Cancel", Left = 262, Top = 150, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnApply;

            Controls.AddRange(new Control[] { lblText, _text, _prefix, _suffix, _at, _pos, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _text;
        }

        private void OnApply(object sender, EventArgs e)
        {
            if (_text.TextLength == 0)
            {
                SetError(_text, "Enter the text to add.");
                return;
            }
            SetError(_text, null);

            string add = _text.Text;
            int pos = (int)_pos.Value;
            bool prefix = _prefix.Checked, suffix = _suffix.Checked;

            bool ok = RunOperation(AddTextCommand.Def, CurrentSelection, ctx =>
            {
                foreach (Excel.Range cell in ctx.Target.Cells)
                {
                    object v = cell.Value2;
                    if (v == null) continue;
                    string s = v.ToString();
                    string result;
                    if (prefix) result = add + s;
                    else if (suffix) result = s + add;
                    else
                    {
                        int i = Math.Min(pos - 1, s.Length);
                        if (i < 0) i = 0;
                        result = s.Insert(i, add);
                    }
                    cell.Value2 = result;
                }
            });

            if (ok) Close();
        }
    }

    /// <summary>Remove numeric, alphabetic, non-printing, or custom characters from the selection.</summary>
    [ExcelCommand]
    public sealed class RemoveCharactersCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "Text.RemoveCharacters",
            Label = "Remove Characters",
            Screentip = "Remove Characters",
            Supertip = "Strip numeric, alphabetic, non-printing or a custom set of characters from every selected cell.",
            ImageId = "RemoveCharacters",
            Tab = "Editing",
            Group = "Text",
            Order = 21,
            Scope = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot
        };

        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new RemoveCharactersDialog(); }
    }

    internal sealed class RemoveCharactersDialog : DialogBase
    {
        private readonly CheckBox _numeric = new CheckBox { Text = "Numeric digits (0-9)" };
        private readonly CheckBox _alpha = new CheckBox { Text = "Letters (A-Z, a-z)" };
        private readonly CheckBox _nonPrint = new CheckBox { Text = "Non-printing / control characters" };
        private readonly CheckBox _space = new CheckBox { Text = "Spaces" };
        private readonly CheckBox _customOn = new CheckBox { Text = "Custom characters:" };
        private readonly TextBox _custom = new TextBox { Enabled = false };

        public RemoveCharactersDialog()
        {
            Text = RemoveCharactersCommand.Def.Label;
            ClientSize = new Size(340, 230);

            _numeric.SetBounds(16, 12, 300, 22);
            _alpha.SetBounds(16, 36, 300, 22);
            _nonPrint.SetBounds(16, 60, 300, 22);
            _space.SetBounds(16, 84, 300, 22);
            _customOn.SetBounds(16, 108, 130, 22);
            _custom.SetBounds(150, 108, 166, 23);
            _customOn.CheckedChanged += (s, e) => _custom.Enabled = _customOn.Checked;

            var apply = new Button { Text = "&Apply", Left = 156, Top = 185, Width = 80 };
            var cancel = new Button { Text = "&Cancel", Left = 242, Top = 185, Width = 80, DialogResult = DialogResult.Cancel };
            apply.Click += OnApply;

            Controls.AddRange(new Control[] { _numeric, _alpha, _nonPrint, _space, _customOn, _custom, apply, cancel });
            WireButtons(apply, cancel);
            ActiveControl = _numeric;
        }

        private void OnApply(object sender, EventArgs e)
        {
            bool num = _numeric.Checked, alpha = _alpha.Checked, np = _nonPrint.Checked, sp = _space.Checked;
            string custom = _customOn.Checked ? _custom.Text : string.Empty;

            if (!num && !alpha && !np && !sp && custom.Length == 0)
            {
                SetError(_customOn, "Choose at least one option.");
                return;
            }
            SetError(_customOn, null);

            var customSet = new System.Collections.Generic.HashSet<char>(custom);

            bool ok = RunOperation(RemoveCharactersCommand.Def, CurrentSelection, ctx =>
            {
                foreach (Excel.Range cell in ctx.Target.Cells)
                {
                    object v = cell.Value2;
                    if (v == null) continue;
                    string s = v.ToString();
                    var sb = new StringBuilder(s.Length);
                    foreach (char c in s)
                    {
                        if (num && char.IsDigit(c)) continue;
                        if (alpha && char.IsLetter(c)) continue;
                        if (np && char.IsControl(c)) continue;
                        if (sp && c == ' ') continue;
                        if (customSet.Contains(c)) continue;
                        sb.Append(c);
                    }
                    cell.Value2 = sb.ToString();
                }
            });

            if (ok) Close();
        }
    }
}
