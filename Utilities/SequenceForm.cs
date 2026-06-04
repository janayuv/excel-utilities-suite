using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using utilities.Services;

namespace utilities
{
    // ── Sequence data model ──────────────────────────────────────────────────

    internal enum SequenceType
    {
        Numbers       = 0,
        RomanNumerals = 1,
        WeekdaysShort = 2,
        WeekdaysFull  = 3,
        MonthsShort   = 4,
        MonthsFull    = 5,
    }

    internal sealed class SequenceDefinition
    {
        public string       Name      { get; set; }
        public SequenceType Type      { get; set; }
        public long         Start     { get; set; }
        public bool         HasEnd    { get; set; }
        public long         End       { get; set; }
        public long         Increment { get; set; }
        public int          Digits    { get; set; }
        public string       Prefix    { get; set; }
        public string       Suffix    { get; set; }
        public long         Current   { get; set; }   // next value to insert

        private static readonly string[] _weekShort = { "Mon","Tue","Wed","Thu","Fri","Sat","Sun" };
        private static readonly string[] _weekFull  = { "Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday" };
        private static readonly string[] _monShort  = { "Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec" };
        private static readonly string[] _monFull   = { "January","February","March","April","May","June","July","August","September","October","November","December" };

        public SequenceDefinition()
        {
            Name = "Sequence"; Start = 1; End = 10; Increment = 1;
            Prefix = string.Empty; Suffix = string.Empty;
        }

        public void ResetCurrent() { Current = Start; }

        public string Format(long value)
        {
            switch (Type)
            {
                case SequenceType.RomanNumerals:
                    return Prefix + ToRoman((int)value) + Suffix;
                case SequenceType.WeekdaysShort:
                    return Prefix + _weekShort[Cycle(value, 7)]  + Suffix;
                case SequenceType.WeekdaysFull:
                    return Prefix + _weekFull [Cycle(value, 7)]  + Suffix;
                case SequenceType.MonthsShort:
                    return Prefix + _monShort [Cycle(value, 12)] + Suffix;
                case SequenceType.MonthsFull:
                    return Prefix + _monFull  [Cycle(value, 12)] + Suffix;
                default:
                    string abs = Math.Abs(value).ToString();
                    if (Digits > 0) abs = abs.PadLeft(Digits, '0');
                    return Prefix + (value < 0 ? "-" + abs : abs) + Suffix;
            }
        }

        private static int Cycle(long v, int len) => (int)(((v - 1) % len + len) % len);

        private static string ToRoman(int n)
        {
            if (n <= 0 || n > 3999) return n.ToString();
            int[]    vals = { 1000,900,500,400,100,90,50,40,10,9,5,4,1 };
            string[] syms = { "M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I" };
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < vals.Length; i++)
                while (n >= vals[i]) { sb.Append(syms[i]); n -= vals[i]; }
            return sb.ToString();
        }

        // ── Persistence (SettingsService key-value store) ─────────────────
        private const string K = "Seq.";

        public static List<SequenceDefinition> LoadAll()
        {
            int count = SettingsService.GetInt(K + "Count", 0);
            var list  = new List<SequenceDefinition>(count);
            for (int i = 0; i < count; i++)
            {
                string p = K + i + ".";
                list.Add(new SequenceDefinition
                {
                    Name      = SettingsService.Get(p + "Name", "Sequence " + (i + 1)),
                    Type      = (SequenceType)SettingsService.GetInt(p + "Type", 0),
                    Start     = ParseLong(SettingsService.Get(p + "Start"),  1),
                    HasEnd    = SettingsService.GetBool(p + "HasEnd", false),
                    End       = ParseLong(SettingsService.Get(p + "End"),    10),
                    Increment = ParseLong(SettingsService.Get(p + "Inc"),    1),
                    Digits    = SettingsService.GetInt(p + "Digits", 0),
                    Prefix    = SettingsService.Get(p + "Prefix", ""),
                    Suffix    = SettingsService.Get(p + "Suffix", ""),
                    Current   = ParseLong(SettingsService.Get(p + "Cur"),    1),
                });
            }
            if (list.Count == 0)
                list.Add(new SequenceDefinition { Name = "Sequence 1" });
            return list;
        }

        public static void SaveAll(IList<SequenceDefinition> list)
        {
            SettingsService.Set(K + "Count", list.Count);
            for (int i = 0; i < list.Count; i++)
            {
                string p = K + i + ".";
                var d = list[i];
                SettingsService.Set(p + "Name",   d.Name   ?? "");
                SettingsService.Set(p + "Type",   (int)d.Type);
                SettingsService.Set(p + "Start",  d.Start.ToString());
                SettingsService.Set(p + "HasEnd", d.HasEnd);
                SettingsService.Set(p + "End",    d.End.ToString());
                SettingsService.Set(p + "Inc",    d.Increment.ToString());
                SettingsService.Set(p + "Digits", d.Digits);
                SettingsService.Set(p + "Prefix", d.Prefix ?? "");
                SettingsService.Set(p + "Suffix", d.Suffix ?? "");
                SettingsService.Set(p + "Cur",    d.Current.ToString());
            }
        }

        private static long ParseLong(string s, long fallback = 0)
        {
            long v; return long.TryParse(s, out v) ? v : fallback;
        }
    }

    // ── Main "Insert Sequence Number" dialog ─────────────────────────────────

    public partial class SequenceForm : Form
    {
        private readonly ComboBox _fillOrder = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ListView _list      = new ListView
        {
            FullRowSelect = true, MultiSelect = false, View = View.Details,
            GridLines = true, HideSelection = false, BorderStyle = BorderStyle.FixedSingle
        };
        private readonly Button _btnNew    = new Button { Text = "+ New" };
        private readonly Button _btnEdit   = new Button { Text = "/ Edit" };
        private readonly Button _btnReset  = new Button { Text = "Reset" };
        private readonly Button _btnRemove = new Button { Text = "Remove" };
        private readonly Button _btnRmAll  = new Button { Text = "Remove all" };
        private readonly Button _btnOK     = new Button { Text = "OK" };
        private readonly Button _btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
        private readonly Button _btnApply  = new Button { Text = "Apply" };

        private List<SequenceDefinition> _seqs;

        public SequenceForm()
        {
            InitializeComponent();

            Text = "Insert Sequence Number";
            ClientSize = new Size(540, 340);
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;
            KeyPreview = true;

            // Fill order row
            var lblFill = new Label { Text = "Fill order:", AutoSize = true };
            lblFill.SetBounds(12, 17, 70, 20);
            _fillOrder.Items.AddRange(new object[] {
                "Fill vertically cell after cell",
                "Fill horizontally cell after cell"
            });
            _fillOrder.SelectedIndex = SettingsService.GetInt("Seq.FillOrder", 0);
            _fillOrder.SetBounds(90, 13, 252, 23);

            // Sequence list view
            _list.SetBounds(12, 46, 512, 192);
            _list.Columns.Add("Name",         120);
            _list.Columns.Add("Type",          90);
            _list.Columns.Add("Start",         60);
            _list.Columns.Add("Increment",     70);
            _list.Columns.Add("Previous",      65);
            _list.Columns.Add("Next",          65);
            _list.DoubleClick          += (s, e) => EditSelected();
            _list.SelectedIndexChanged += (s, e) => UpdateButtons();

            // Toolbar buttons
            const int ty = 246;
            SetupBtn(_btnNew,    12,  ty, 68);
            SetupBtn(_btnEdit,   84,  ty, 58);
            SetupBtn(_btnReset,  146, ty, 58);
            SetupBtn(_btnRemove, 208, ty, 68);
            SetupBtn(_btnRmAll,  280, ty, 80);

            _btnNew.Click    += (s, e) => AddNew();
            _btnEdit.Click   += (s, e) => EditSelected();
            _btnReset.Click  += (s, e) => ResetSelected();
            _btnRemove.Click += (s, e) => RemoveSelected();
            _btnRmAll.Click  += (s, e) =>
            {
                if (Confirm("Remove all sequences?")) { _seqs.Clear(); RefreshList(); }
            };

            // OK / Apply / Cancel
            _btnApply.SetBounds(264, 298, 75, 28);
            _btnOK.SetBounds(350, 298, 75, 28);
            _btnCancel.SetBounds(438, 298, 75, 28);

            _btnOK.Click    += (s, e) => { if (DoApply()) { DialogResult = DialogResult.OK; Close(); } };
            _btnApply.Click += (s, e) => DoApply();

            Controls.AddRange(new Control[] {
                lblFill, _fillOrder, _list,
                _btnNew, _btnEdit, _btnReset, _btnRemove, _btnRmAll,
                _btnApply, _btnOK, _btnCancel
            });
            AcceptButton = _btnOK;
            CancelButton = _btnCancel;

            _seqs = SequenceDefinition.LoadAll();
            RefreshList();
        }

        private static void SetupBtn(Button btn, int x, int y, int w)
        {
            btn.SetBounds(x, y, w, 24);
            btn.FlatStyle = FlatStyle.System;
        }

        private void RefreshList()
        {
            int selIdx = _list.SelectedIndices.Count > 0 ? _list.SelectedIndices[0] : 0;
            _list.Items.Clear();
            foreach (var d in _seqs)
            {
                long prev = d.Current == d.Start ? d.Start - d.Increment : d.Current - d.Increment;
                var item = new ListViewItem(d.Name ?? "");
                item.SubItems.Add(d.Type.ToString());
                item.SubItems.Add(d.Start.ToString());
                item.SubItems.Add(d.Increment.ToString());
                item.SubItems.Add(d.Format(prev));
                item.SubItems.Add(d.Format(d.Current));
                _list.Items.Add(item);
            }
            if (_list.Items.Count > 0)
            {
                int idx = Math.Min(selIdx, _list.Items.Count - 1);
                _list.Items[idx].Selected = true;
                _list.EnsureVisible(idx);
            }
            UpdateButtons();
        }

        private void UpdateButtons()
        {
            bool has = _list.SelectedIndices.Count > 0;
            _btnEdit.Enabled = _btnReset.Enabled = _btnRemove.Enabled =
                _btnOK.Enabled = _btnApply.Enabled = has;
        }

        private SequenceDefinition Selected
        {
            get
            {
                if (_list.SelectedIndices.Count == 0) return null;
                int i = _list.SelectedIndices[0];
                return i < _seqs.Count ? _seqs[i] : null;
            }
        }

        private void AddNew()
        {
            var def = new SequenceDefinition { Name = "Sequence " + (_seqs.Count + 1) };
            using (var dlg = new CreateSequenceDialog(def, _seqs))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                _seqs.Add(dlg.Result);
                RefreshList();
                if (_seqs.Count > 0)
                {
                    _list.Items[_seqs.Count - 1].Selected = true;
                    _list.EnsureVisible(_seqs.Count - 1);
                }
            }
        }

        private void EditSelected()
        {
            var sel = Selected;
            if (sel == null) return;
            int idx = _list.SelectedIndices[0];
            using (var dlg = new CreateSequenceDialog(sel, _seqs))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                _seqs[idx] = dlg.Result;
                RefreshList();
            }
        }

        private void ResetSelected()
        {
            var sel = Selected;
            if (sel == null) return;
            sel.ResetCurrent();
            RefreshList();
        }

        private void RemoveSelected()
        {
            var sel = Selected;
            if (sel == null) return;
            if (!Confirm("Remove sequence '" + sel.Name + "'?")) return;
            _seqs.RemoveAt(_list.SelectedIndices[0]);
            RefreshList();
        }

        private bool DoApply()
        {
            var sel = Selected;
            if (sel == null) { MessageBox.Show("Select a sequence to insert.", Text); return false; }

            var app = Globals.ThisAddIn.Application;
            var rng = app.Selection as Excel.Range;
            if (rng == null) { MessageBox.Show("Please select cells to fill.", Text); return false; }

            bool vertical = _fillOrder.SelectedIndex == 0;
            int rows = rng.Rows.Count, cols = rng.Columns.Count;
            long cur = sel.Current;
            bool stop = false;

            for (int outer = 1; outer <= (vertical ? cols : rows) && !stop; outer++)
            {
                for (int inner = 1; inner <= (vertical ? rows : cols) && !stop; inner++)
                {
                    if (sel.HasEnd && (sel.Increment >= 0 ? cur > sel.End : cur < sel.End))
                    {
                        stop = true; break;
                    }
                    int r = vertical ? inner : outer;
                    int c = vertical ? outer : inner;
                    var cell = (Excel.Range)rng.Cells[r, c];

                    // Skip filtered/hidden cells so the visible cells get a consecutive
                    // run (55,56,57…) instead of values being wasted on hidden rows.
                    if (cell.EntireRow.Hidden || cell.EntireColumn.Hidden)
                        continue;

                    cell.Value = sel.Format(cur);
                    cur += sel.Increment;
                }
            }

            sel.Current = cur;
            SettingsService.Set("Seq.FillOrder", _fillOrder.SelectedIndex);
            SequenceDefinition.SaveAll(_seqs);
            RefreshList();
            return true;
        }

        private static bool Confirm(string msg) =>
            MessageBox.Show(msg, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                == DialogResult.Yes;

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            { DialogResult = DialogResult.Cancel; Close(); e.Handled = true; return; }
            base.OnKeyDown(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            SequenceDefinition.SaveAll(_seqs);
        }
    }

    // ── "Create / Edit Sequence Number" sub-dialog ───────────────────────────

    internal sealed class CreateSequenceDialog : Form
    {
        private readonly TextBox       _name   = new TextBox();
        private readonly ComboBox      _type   = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly NumericUpDown _start  = new NumericUpDown { Minimum = -999999999, Maximum = 999999999 };
        private readonly CheckBox      _hasEnd = new CheckBox { Text = "End number:" };
        private readonly NumericUpDown _end    = new NumericUpDown { Minimum = -999999999, Maximum = 999999999, Value = 10 };
        private readonly NumericUpDown _inc    = new NumericUpDown { Minimum = -999999, Maximum = 999999, Value = 1 };
        private readonly NumericUpDown _digits = new NumericUpDown { Minimum = 0, Maximum = 20, Value = 0 };
        private readonly TextBox       _prefix = new TextBox();
        private readonly TextBox       _suffix = new TextBox();
        private readonly ListBox       _preview = new ListBox();

        private readonly List<SequenceDefinition> _all;
        private readonly string _origName;

        public SequenceDefinition Result { get; private set; }

        public CreateSequenceDialog(SequenceDefinition src, List<SequenceDefinition> allSeqs)
        {
            _all      = allSeqs;
            _origName = src.Name ?? "";

            Text = "Create / Edit Sequence";
            ClientSize = new Size(490, 475);
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            ShowInTaskbar = false; StartPosition = FormStartPosition.CenterParent; KeyPreview = true;

            _type.Items.AddRange(new object[] {
                "Numbers", "Roman Numerals",
                "Weekdays (Mon–Sun)", "Weekdays (Monday–Sunday)",
                "Months (Jan–Dec)",   "Months (January–December)"
            });

            _name.Text      = src.Name ?? "";
            _type.SelectedIndex = (int)src.Type;
            _start.Value    = Clamp(src.Start);
            _hasEnd.Checked = src.HasEnd;
            _end.Value      = Clamp(src.End);
            _end.Enabled    = src.HasEnd;
            _inc.Value      = Clamp(src.Increment, -999999, 999999);
            _digits.Value   = Math.Max(0, Math.Min(20, src.Digits));
            _prefix.Text    = src.Prefix ?? "";
            _suffix.Text    = src.Suffix ?? "";

            int lx = 14, cx = 158, cw = 120, y = 18;
            var hdr = new Label { Text = "Sequence", Left = lx, Top = y, AutoSize = true };
            hdr.Font = new Font(hdr.Font, FontStyle.Bold);
            y += 26;

            AddFieldRow("Name:",              ref y, lx, cx, cw, _name);
            AddFieldRow("Type:",              ref y, lx, cx, 170, _type);
            AddFieldRow("Start:",             ref y, lx, cx, cw, _start);
            _hasEnd.SetBounds(lx, y + 2, 140, 20);
            _end.SetBounds(cx, y, cw, 23);
            y += 30;
            AddFieldRow("Increment:",         ref y, lx, cx, cw, _inc);
            AddFieldRow("No. of digits:",     ref y, lx, cx, cw, _digits);
            AddFieldRow("Prefix (optional):", ref y, lx, cx, cw, _prefix);
            AddFieldRow("Suffix (optional):", ref y, lx, cx, cw, _suffix);

            var lblPrev = new Label { Text = "Preview", Left = 300, Top = 18, AutoSize = true };
            lblPrev.Font = new Font(lblPrev.Font, FontStyle.Bold);
            _preview.SetBounds(300, 42, 175, 388);
            _preview.Font = new Font("Courier New", 9f);
            _preview.BorderStyle = BorderStyle.FixedSingle;

            var btnOK     = new Button { Text = "OK",     Width = 75 };
            var btnCancel = new Button { Text = "Cancel", Width = 75, DialogResult = DialogResult.Cancel };
            btnOK.SetBounds(300, 438, 75, 28); btnCancel.SetBounds(390, 438, 75, 28);
            btnOK.Click += OnOK;

            Controls.AddRange(new Control[] { hdr, _hasEnd, _end, lblPrev, _preview, btnOK, btnCancel });
            AcceptButton = btnOK; CancelButton = btnCancel;

            // Wire live-preview — type change also enables/disables digits
            _type.SelectedIndexChanged += (s, e) => { _digits.Enabled = _type.SelectedIndex == 0; UpdatePreview(); };
            _digits.Enabled = _type.SelectedIndex == 0;
            _name.TextChanged      += (s, e) => UpdatePreview();
            _start.ValueChanged    += (s, e) => UpdatePreview();
            _end.ValueChanged      += (s, e) => UpdatePreview();
            _inc.ValueChanged      += (s, e) => UpdatePreview();
            _digits.ValueChanged   += (s, e) => UpdatePreview();
            _prefix.TextChanged    += (s, e) => UpdatePreview();
            _suffix.TextChanged    += (s, e) => UpdatePreview();
            _hasEnd.CheckedChanged += (s, e) => { _end.Enabled = _hasEnd.Checked; UpdatePreview(); };

            UpdatePreview();
        }

        private static decimal Clamp(long v, long min = -999999999, long max = 999999999)
            => v < min ? (decimal)min : v > max ? (decimal)max : (decimal)v;

        private void AddFieldRow(string label, ref int y, int lx, int cx, int cw, Control ctrl)
        {
            Controls.Add(new Label { Text = label, Left = lx, Top = y + 3, AutoSize = true });
            ctrl.SetBounds(cx, y, cw, 23);
            Controls.Add(ctrl);
            y += 30;
        }

        private void UpdatePreview()
        {
            _preview.Items.Clear();
            try
            {
                var type  = (SequenceType)_type.SelectedIndex;
                long start = (long)_start.Value;
                long inc   = (long)_inc.Value;
                bool hasE  = _hasEnd.Checked;
                long end   = (long)_end.Value;

                if (inc == 0) { _preview.Items.Add("(increment is 0)"); return; }

                var tmp = new SequenceDefinition
                {
                    Type = type, Start = start, Increment = inc,
                    Digits = (int)_digits.Value, Prefix = _prefix.Text, Suffix = _suffix.Text
                };
                long cur = start;
                for (int i = 0; i < 20; i++)
                {
                    if (hasE && (inc > 0 ? cur > end : cur < end)) break;
                    _preview.Items.Add(tmp.Format(cur));
                    cur += inc;
                }
            }
            catch { }
        }

        private void OnOK(object sender, EventArgs e)
        {
            string nm = _name.Text.Trim();
            if (string.IsNullOrEmpty(nm)) { MessageBox.Show("Enter a name.", Text); return; }
            if (_all.Any(d => string.Equals(d.Name, nm, StringComparison.OrdinalIgnoreCase)
                           && !string.Equals(d.Name, _origName, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show("A sequence named '" + nm + "' already exists.", Text); return; }

            long start = (long)_start.Value;
            Result = new SequenceDefinition
            {
                Name      = nm,
                Type      = (SequenceType)_type.SelectedIndex,
                Start     = start,
                HasEnd    = _hasEnd.Checked,
                End       = (long)_end.Value,
                Increment = (long)_inc.Value,
                Digits    = (int)_digits.Value,
                Prefix    = _prefix.Text,
                Suffix    = _suffix.Text,
                Current   = start
            };
            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            { DialogResult = DialogResult.Cancel; Close(); e.Handled = true; return; }
            base.OnKeyDown(e);
        }
    }
}
