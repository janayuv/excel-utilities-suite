using System;
using System.Drawing;
using System.Windows.Forms;
using utilities.Dialogs;
using utilities.Services;

namespace utilities.Commands.Tools
{
    [ExcelCommand]
    public sealed class GridFocusCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id = "View.GridFocus", Label = "Grid Focus",
            Screentip = "Grid Focus",
            Supertip = "Highlight the row and column of the active cell for easier reading.",
            ImageId = "GridFocus",
            Tab = "Utilities", Group = "View", Order = 5,
            RequiresSelection = false, UndoMode = UndoMode.None
        };
        public override CommandDefinition Definition { get { return Def; } }
        protected override DialogBase CreateDialog() { return new GridFocusDialog(); }
    }

    internal sealed class GridFocusDialog : DialogBase
    {
        private readonly RadioButton _rbCriss  = new RadioButton { Text = "Criss-cross",  Checked = true, AutoSize = true };
        private readonly RadioButton _rbStr    = new RadioButton { Text = "Straight line", AutoSize = true };
        private readonly RadioButton _rbVert   = new RadioButton { Text = "Vertical line", AutoSize = true };
        private readonly RadioButton _rbWide   = new RadioButton { Text = "Wide Stripe",   Checked = true, AutoSize = true };
        private readonly RadioButton _rbHollow = new RadioButton { Text = "Hollow Type",   AutoSize = true };
        private readonly RadioButton _rbLine   = new RadioButton { Text = "Line Type",     AutoSize = true };

        private readonly Button   _colorBtn = new Button  { Width = 80, Height = 22 };
        private readonly Label    _colorHex = new Label   { AutoSize = true };
        private readonly TrackBar _thick    = new TrackBar { Minimum = 0, Maximum = 2, Value = 1, TickFrequency = 1, Width = 110, Height = 30 };
        private readonly Label    _thickLbl = new Label   { AutoSize = true };
        private readonly TrackBar _trans    = new TrackBar { Minimum = 0, Maximum = 100, Value = 70, TickFrequency = 10, Width = 110, Height = 30 };
        private readonly Label    _transLbl = new Label   { AutoSize = true };

        private readonly CheckBox _chkArea = new CheckBox { Text = "Highlight the area of active cell",     AutoSize = true };
        private readonly CheckBox _chkEdit = new CheckBox { Text = "Highlight the area on editing of cell", AutoSize = true };
        private readonly CheckBox _chkOn   = new CheckBox { Text = "Enable Grid Focus",                     AutoSize = true };

        private readonly Panel _preview = new Panel { BorderStyle = BorderStyle.FixedSingle };
        private Color _color = Color.FromArgb(0x5F, 0xC8, 0xD8);

        public GridFocusDialog()
        {
            Text       = "Grid Focus Settings";
            ClientSize = new Size(634, 416);
            LoadSettings();
            BuildLayout();
        }

        private void LoadSettings()
        {
            var s = GridFocusService.Settings;
            _rbCriss.Checked  = s.Shape == GridFocusShape.CrissCross;
            _rbStr.Checked    = s.Shape == GridFocusShape.StraightLine;
            _rbVert.Checked   = s.Shape == GridFocusShape.VerticalLine;
            _rbWide.Checked   = s.Style == GridFocusStyle.WideStripe;
            _rbHollow.Checked = s.Style == GridFocusStyle.HollowType;
            _rbLine.Checked   = s.Style == GridFocusStyle.LineType;
            _color            = s.Color;
            _thick.Value      = s.Thickness;
            _trans.Value      = s.Transparency;
            _chkArea.Checked  = s.HighlightActiveArea;
            _chkEdit.Checked  = s.HighlightOnEditing;
            _chkOn.Checked    = s.Enabled;
        }

        private void BuildLayout()
        {
            var grpShape = new GroupBox { Text = "Shapes", Left = 12, Top = 10, Width = 130, Height = 96 };
            _rbCriss.SetBounds(8, 16, 115, 22); _rbStr.SetBounds(8, 40, 115, 22); _rbVert.SetBounds(8, 64, 115, 22);
            grpShape.Controls.AddRange(new Control[] { _rbCriss, _rbStr, _rbVert });

            var grpStyle = new GroupBox { Text = "Styles", Left = 152, Top = 10, Width = 120, Height = 96 };
            _rbWide.SetBounds(8, 16, 108, 22); _rbHollow.SetBounds(8, 40, 108, 22); _rbLine.SetBounds(8, 64, 108, 22);
            grpStyle.Controls.AddRange(new Control[] { _rbWide, _rbHollow, _rbLine });

            var grpOpts = new GroupBox { Text = "Options", Left = 12, Top = 114, Width = 258, Height = 108 };
            var lblC = new Label { Text = "Color:",        Left = 8, Top = 20, AutoSize = true };
            _colorBtn.SetBounds(68, 18, 80, 22); _colorHex.SetBounds(155, 20, 80, 16);
            _colorBtn.Click += OnColorClick;
            var lblT = new Label { Text = "Thickness:",    Left = 8, Top = 52, AutoSize = true };
            _thick.SetBounds(80, 48, 110, 30); _thickLbl.SetBounds(196, 52, 50, 16);
            _thick.ValueChanged += (s, e) => { SyncThickLabel(); _preview.Invalidate(); };
            var lblP = new Label { Text = "Transparency:", Left = 8, Top = 84, AutoSize = true };
            _trans.SetBounds(96, 80, 110, 30); _transLbl.SetBounds(212, 84, 45, 16);
            _trans.ValueChanged += (s, e) => { SyncTransLabel(); _preview.Invalidate(); };
            grpOpts.Controls.AddRange(new Control[] { lblC, _colorBtn, _colorHex, lblT, _thick, _thickLbl, lblP, _trans, _transLbl });

            _chkArea.SetBounds(12, 230, 270, 22);
            _chkEdit.SetBounds(12, 254, 270, 22);
            _chkOn.SetBounds(12,  278, 270, 22);

            var lblPrev = new Label { Text = "Preview", Left = 290, Top = 10, AutoSize = true };
            _preview.SetBounds(290, 28, 330, 280);
            _preview.Paint += OnPreviewPaint;

            var btnDef = new Button { Text = "Default Settings", Left = 12,  Top = 378, Width = 110 };
            var ok     = new Button { Text = "OK",     Left = 444, Top = 378, Width = 80, DialogResult = DialogResult.None };
            var cancel = new Button { Text = "Cancel", Left = 534, Top = 378, Width = 80, DialogResult = DialogResult.Cancel };
            btnDef.Click += OnDefault;
            ok.Click     += OnOk;

            foreach (var rb in new RadioButton[] { _rbCriss, _rbStr, _rbVert, _rbWide, _rbHollow, _rbLine })
                rb.CheckedChanged += (s, e) => { SyncThickEnabled(); _preview.Invalidate(); };

            Controls.AddRange(new Control[] { grpShape, grpStyle, grpOpts, _chkArea, _chkEdit, _chkOn, lblPrev, _preview, btnDef, ok, cancel });
            WireButtons(ok, cancel);
            SyncThickLabel(); SyncTransLabel(); SyncColorBtn(); SyncThickEnabled();
        }

        private void SyncThickLabel()   => _thickLbl.Text = _thick.Value == 0 ? "Thin" : _thick.Value == 1 ? "Mid" : "Thick";
        private void SyncTransLabel()   => _transLbl.Text = _trans.Value + "%";
        private void SyncThickEnabled() { bool en = !_rbWide.Checked; _thick.Enabled = en; _thickLbl.Enabled = en; }
        private void SyncColorBtn()
        {
            _colorBtn.BackColor = _color;
            _colorHex.Text = "#" + _color.R.ToString("X2") + _color.G.ToString("X2") + _color.B.ToString("X2");
        }

        private void OnColorClick(object sender, EventArgs e)
        {
            using (var dlg = new ColorDialog { Color = _color, FullOpen = true })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                _color = dlg.Color;
                SyncColorBtn();
                _preview.Invalidate();
            }
        }

        private void OnDefault(object sender, EventArgs e)
        {
            _rbCriss.Checked = true; _rbWide.Checked = true;
            _color = Color.FromArgb(0x5F, 0xC8, 0xD8);
            _thick.Value = 1; _trans.Value = 70;
            _chkArea.Checked = false; _chkEdit.Checked = false; _chkOn.Checked = false;
            SyncColorBtn(); SyncThickLabel(); SyncTransLabel(); SyncThickEnabled();
            _preview.Invalidate();
        }

        private void OnOk(object sender, EventArgs e)
        {
            var s = GridFocusService.Settings;
            s.Shape  = _rbCriss.Checked ? GridFocusShape.CrissCross :
                       _rbStr.Checked   ? GridFocusShape.StraightLine : GridFocusShape.VerticalLine;
            s.Style  = _rbWide.Checked  ? GridFocusStyle.WideStripe :
                       _rbHollow.Checked? GridFocusStyle.HollowType  : GridFocusStyle.LineType;
            s.Color              = _color;
            s.Thickness          = _thick.Value;
            s.Transparency       = _trans.Value;
            s.HighlightActiveArea = _chkArea.Checked;
            s.HighlightOnEditing  = _chkEdit.Checked;
            s.Enabled            = _chkOn.Checked;
            GridFocusService.OnSelectionChange(Globals.ThisAddIn.Application);
            Close();
        }

        private void OnPreviewPaint(object sender, PaintEventArgs e)
        {
            Graphics  g    = e.Graphics;
            Rectangle rect = _preview.ClientRectangle;
            rect.Inflate(-6, -6);

            const int cols = 7, rows = 8, aC = 4, aR = 3;
            int cw = rect.Width / cols, ch = rect.Height / rows;
            int gX = rect.X + (rect.Width  - cw * cols) / 2;
            int gY = rect.Y + (rect.Height - ch * rows) / 2;
            int gW = cw * cols, gH = ch * rows;

            g.FillRectangle(Brushes.White, gX, gY, gW, gH);

            int   alpha  = Math.Max(10, (int)((100 - _trans.Value) / 100.0 * 220));
            Color hc     = Color.FromArgb(alpha, _color);
            int   lw     = _thick.Value == 0 ? 1 : _thick.Value == 1 ? 2 : 3;

            GridFocusShape sh = _rbCriss.Checked ? GridFocusShape.CrissCross :
                                _rbStr.Checked   ? GridFocusShape.StraightLine : GridFocusShape.VerticalLine;
            GridFocusStyle st = _rbWide.Checked  ? GridFocusStyle.WideStripe :
                                _rbHollow.Checked? GridFocusStyle.HollowType  : GridFocusStyle.LineType;

            if (sh != GridFocusShape.VerticalLine)
                PaintBand(g, new Rectangle(gX, gY + aR * ch, gW, ch), hc, st, lw, isRow: true);
            if (sh != GridFocusShape.StraightLine)
                PaintBand(g, new Rectangle(gX + aC * cw, gY, cw, gH), hc, st, lw, isRow: false);

            using (var pen = new Pen(Color.Silver))
            {
                for (int c = 0; c <= cols; c++) g.DrawLine(pen, gX + c * cw, gY, gX + c * cw, gY + gH);
                for (int r = 0; r <= rows; r++) g.DrawLine(pen, gX, gY + r * ch, gX + gW, gY + r * ch);
            }
            g.DrawRectangle(Pens.Black, gX + aC * cw, gY + aR * ch, cw, ch);
        }

        private static void PaintBand(Graphics g, Rectangle rect, Color color, GridFocusStyle style, int lw, bool isRow)
        {
            switch (style)
            {
                case GridFocusStyle.WideStripe:
                    using (var b = new SolidBrush(color)) g.FillRectangle(b, rect); break;
                case GridFocusStyle.HollowType:
                    using (var p = new Pen(color, lw)) g.DrawRectangle(p, rect); break;
                case GridFocusStyle.LineType:
                    var lr = isRow ? new Rectangle(rect.X, rect.Y, rect.Width, lw)
                                   : new Rectangle(rect.X, rect.Y, lw, rect.Height);
                    using (var b = new SolidBrush(color)) g.FillRectangle(b, lr); break;
            }
        }
    }
}
