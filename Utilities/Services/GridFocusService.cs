using System.Collections.Generic;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using MSO   = Microsoft.Office.Core;

namespace utilities.Services
{
    public enum GridFocusShape { CrissCross, StraightLine, VerticalLine }
    public enum GridFocusStyle { WideStripe, HollowType, LineType }

    public sealed class GridFocusSettings
    {
        public GridFocusShape Shape             { get; set; } = GridFocusShape.CrissCross;
        public GridFocusStyle Style             { get; set; } = GridFocusStyle.WideStripe;
        public Color          Color             { get; set; } = Color.FromArgb(0x5F, 0xC8, 0xD8);
        public int            Thickness         { get; set; } = 1;  // 0=Thin 1=Mid 2=Thick
        public int            Transparency      { get; set; } = 70; // percent 0-100
        public bool           HighlightActiveArea  { get; set; } = false;
        public bool           HighlightOnEditing   { get; set; } = false;
        public bool           Enabled           { get; set; } = false;
    }

    /// <summary>
    /// Manages transparent shape overlays that highlight the active cell's
    /// row and/or column. Called from ThisAddIn.OnSheetSelectionChange.
    /// </summary>
    public static class GridFocusService
    {
        private const string NameRow = "__GF_Row__";
        private const string NameCol = "__GF_Col__";

        public static readonly GridFocusSettings Settings = new GridFocusSettings();

        private static Excel.Worksheet _lastSheet;

        /// <summary>Call on every SheetSelectionChange to move the highlight.</summary>
        public static void OnSelectionChange(Excel.Application app)
        {
            try
            {
                Excel.Worksheet ws   = app.ActiveSheet as Excel.Worksheet;
                Excel.Range     cell = app.Selection   as Excel.Range;
                if (ws == null) return;

                if (_lastSheet != null && !ReferenceEquals(_lastSheet, ws))
                    ClearShapes(_lastSheet);

                _lastSheet = ws;
                ClearShapes(ws);

                if (!Settings.Enabled || cell == null) return;
                AddHighlight(ws, cell);
            }
            catch { /* never crash Excel's selection change pipeline */ }
        }

        /// <summary>
        /// Drop the cached sheet reference. Call when a workbook closes so the COM
        /// wrapper is not held past the workbook's lifetime.
        /// </summary>
        public static void ClearLastSheet() { _lastSheet = null; }

        /// <summary>Remove all Grid Focus shapes from every sheet in the workbook.</summary>
        public static void ClearAll(Excel.Application app)
        {
            try
            {
                foreach (Excel.Worksheet ws in app.ActiveWorkbook.Worksheets)
                    ClearShapes(ws);
            }
            catch { }
        }

        // ── Private helpers ───────────────────────────────────────────────────

        private static void ClearShapes(Excel.Worksheet ws)
        {
            var del = new List<Excel.Shape>();
            try
            {
                foreach (Excel.Shape s in ws.Shapes)
                    if (s.Name == NameRow || s.Name == NameCol) del.Add(s);
            }
            catch { }
            foreach (var s in del) try { s.Delete(); } catch { }
        }

        private static void AddHighlight(Excel.Worksheet ws, Excel.Range cell)
        {
            const float Max = 9000f;

            float rowTop    = (float)(double)((Excel.Range)ws.Cells[cell.Row, 1]).Top;
            float rowHeight = (float)(double)((Excel.Range)ws.Rows[cell.Row]).Height;
            float colLeft   = (float)(double)((Excel.Range)ws.Cells[1, cell.Column]).Left;
            float colWidth  = (float)(double)((Excel.Range)ws.Columns[cell.Column]).Width;

            bool doRow = Settings.Shape != GridFocusShape.VerticalLine;
            bool doCol = Settings.Shape != GridFocusShape.StraightLine;

            if (doRow) PlaceShape(ws, NameRow, 0f,      rowTop, Max,      rowHeight, isRow: true);
            if (doCol) PlaceShape(ws, NameCol, colLeft, 0f,     colWidth, Max,       isRow: false);
        }

        private static void PlaceShape(Excel.Worksheet ws, string name,
            float left, float top, float width, float height, bool isRow)
        {
            var   cfg   = Settings;
            float lw    = cfg.Thickness == 0 ? 1f : cfg.Thickness == 1 ? 2f : 4f;
            float actW  = width, actH = height;

            if (cfg.Style == GridFocusStyle.LineType)
            {
                if (isRow) actH = lw;
                else       actW = lw;
            }

            Excel.Shape shp = ws.Shapes.AddShape(
                MSO.MsoAutoShapeType.msoShapeRectangle,
                left, top, actW, actH);

            shp.Name      = name;
            shp.Locked    = false;
            shp.Placement = Excel.XlPlacement.xlFreeFloating;
            shp.ZOrder(MSO.MsoZOrderCmd.msoSendToBack);

            int   rgb   = cfg.Color.R | (cfg.Color.G << 8) | (cfg.Color.B << 16);
            float trans = cfg.Transparency / 100f;

            if (cfg.Style == GridFocusStyle.HollowType)
            {
                shp.Fill.Visible       = MSO.MsoTriState.msoFalse;
                shp.Line.Visible       = MSO.MsoTriState.msoCTrue;
                shp.Line.ForeColor.RGB = rgb;
                shp.Line.Weight        = lw;
                shp.Line.Transparency  = trans;
            }
            else
            {
                shp.Fill.Visible       = MSO.MsoTriState.msoCTrue;
                shp.Fill.ForeColor.RGB = rgb;
                shp.Fill.Transparency  = trans;
                shp.Line.Visible       = MSO.MsoTriState.msoFalse;
            }
        }
    }
}
