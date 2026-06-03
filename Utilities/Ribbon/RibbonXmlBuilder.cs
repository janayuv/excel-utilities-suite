using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using utilities.Commands;

namespace utilities.Ribbon
{
    /// <summary>
    /// Builds the customUI XML from the command registry. Tabs/groups/buttons and every
    /// tooltip come straight from <see cref="CommandDefinition"/>, so the ribbon cannot
    /// drift from the commands. A trailing "Suite" tab hosts the built-in Undo/Help controls.
    /// </summary>
    internal static class RibbonXmlBuilder
    {
        private const string Ns = "http://schemas.microsoft.com/office/2009/07/customui";

        // Preferred left-to-right tab order; unknown tabs are appended alphabetically.
        private static readonly string[] TabOrder =
        {
            "Utilities", "Utilities +"
        };

        // Collapse all feature tabs into two Kutools-style tabs.
        private static readonly Dictionary<string, string> TabRemap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Editing",               "Utilities" },
            { "Insert",                "Utilities" },
            { "Select & Navigate",     "Utilities" },
            { "Formula & Statistics",  "Utilities +" },
            { "Data & Cleaning",       "Utilities +" },
            { "Export / Import",       "Utilities +" },
            { "Workbook & Sheets",     "Utilities +" },
            { "Printing",              "Utilities +" },
        };

        // Fallback imageMso for tools that have no custom PNG embedded resource.
        // All values verified against the Office 2010+ imageMso gallery.
        // The 7 ImageIds with real embedded PNGs are intentionally absent —
        // those continue to use getImage/IconProvider.
        private static readonly Dictionary<string, string> ImageIdToMso = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // ── Color ─────────────────────────────────────────────────────────
            { "SelectByColor",           "SelectCurrentRegion" },
            { "SumCountByColor",         "AutoSum" },
            { "CountByColor",            "AutoSum" },
            // ── Format & Convert ──────────────────────────────────────────────
            { "ClearFormatting",         "ClearFormats" },
            { "FormulasToValues",        "PasteValues" },
            { "RoundValues",             "DecimalDecrease" },
            { "ChangeSign",              "NumberFormatDialog" },
            { "NumbersToWords",          "Spelling" },
            { "CopyCellFormatting",      "FormatPainter" },
            { "AlternateRows",           "FormatCellsDialog" },
            { "ConvertDateFormat",       "FunctionWizard" },
            { "SwapRanges",              "Cut" },
            { "CurrencyFormat",          "NumberFormatDialog" },
            { "UnitConversion",          "NumberFormatDialog" },
            // ── Data ──────────────────────────────────────────────────────────
            { "RemoveDuplicateRows",     "RemoveDuplicates" },
            { "WrapIferror",             "FormulaErrorChecking" },
            { "ToggleRefStyle",          "FormulaBar" },
            { "SplitIntoRows",           "TextToColumns" },
            { "DetectDataTypes",         "DataValidation" },
            { "FuzzyDedupe",             "RemoveDuplicates" },
            { "CalculateAge",            "FunctionWizard" },
            { "UnmergeAndFill",          "UnMergeCell" },
            { "CompareRanges",           "GoToSpecial" },
            { "MergeCellsKeepData",      "MergeCenter" },
            { "HighlightDuplicates",     "ConditionalFormattingHighlightCellsRulesMenu" },
            { "ClearCondFormatting",     "ConditionalFormattingClearRulesMenu" },
            // ── Insert / Fill ─────────────────────────────────────────────────
            { "FillDownBlanks",          "FillDown" },
            { "TransposeRange",          "PasteSpecial" },
            { "CombineRows",             "MergeCenter" },
            { "InsertRandomData",        "FunctionWizard" },
            { "InsertBullets",           "InsertTextBox" },
            { "InsertDateSequence",      "FillDown" },
            { "InsertBlankRows",         "InsertRows" },
            // ── Range / Select ────────────────────────────────────────────────
            { "DeleteBlankRows",         "DeleteRows" },
            { "SelectBlankCells",        "GoToSpecial" },
            { "SelectNonBlankCells",     "GoToSpecial" },
            { "SelectErrorCells",        "FormulaErrorChecking" },
            { "SelectDuplicates",        "RemoveDuplicates" },
            { "SelectUniques",           "ConditionalFormattingHighlightCellsRulesMenu" },
            { "SelectByValue",           "AutoFilter" },
            { "SelectByFontColor",       "FontColor" },
            { "SelectMaxCell",           "ConditionalFormattingTopBottomRulesMenu" },
            { "SelectMinCell",           "ConditionalFormattingTopBottomRulesMenu" },
            { "CopyVisibleCells",        "Copy" },
            { "SelectFirstCell",         "GoToSpecial" },
            { "SelectLastCell",          "GoToSpecial" },
            // ── Text ──────────────────────────────────────────────────────────
            { "AddText",                 "InsertTextBox" },
            { "RemoveCharacters",        "ClearContents" },
            { "ReverseText",             "SortDescendingA" },
            { "ExtractNumbers",          "NumberFormatDialog" },
            { "ExtractText",             "FindAndReplace" },
            { "SplitNames",              "TextToColumns" },
            { "CountWords",              "Spelling" },
            { "AddLeadingZeros",         "NumberFormatDialog" },
            { "RemoveApostrophes",       "ClearContents" },
            { "ProperCase",              "Bold" },
            { "AddLineBreak",            "WrapText" },
            { "SuperSubscript",          "Subscript" },
            // ── Find / Navigate / Sheet ───────────────────────────────────────
            { "FindReplaceSheets",       "FindAndReplace" },
            { "FlipVertical",            "RotateLeft90" },
            { "FlipHorizontal",          "RotateRight90" },
            { "SheetTOC",                "InsertHyperlink" },
            { "SortSheets",              "SortAscendingA" },
            { "RenameSheets",            "NameManager" },
            { "SheetsToFiles",           "SaveAs" },
            { "UnhideAllSheets",         "GroupShowDetail" },
            { "BatchRenameSheets",       "NameManager" },
            { "RefreshPivots",           "RefreshAll" },
            { "ClearHyperlinks",         "HyperlinkRemove" },
            { "AutoFitAll",              "ColumnWidthAutoFit" },
            { "FreezePanes",             "FreezePanes" },
            // ── Export / Import / Workbook ────────────────────────────────────
            { "ExportToCsv",             "SaveAs" },
            { "ExportToJson",            "SaveAs" },
            { "InsertFileList",          "InsertHyperlink" },
            { "ImportFilenames",         "FileOpen" },
            { "CopySheets",              "Copy" },
            { "MergeWorkbooks",          "SaveAs" },
            { "SplitByColumn",           "TextToColumns" },
            { "ExportRangeImage",        "FileOpen" },
            // ── Printing ──────────────────────────────────────────────────────
            { "PrintMultipleWorkbooks",  "FilePrint" },
            { "PrintMultipleSelections", "PrintPreviewAndPrint" },
            { "PrintFirstPage",          "FilePrint" },
            { "PrintReverseOrder",       "SortDescendingA" },
            { "PrintCurrentPage",        "FilePrint" },
            { "PrintSpecifiedPages",     "FilePrint" },
            { "PrintCircleInvalidData",  "DataValidation" },
            { "PrintChartsOnly",         "ChartInsert" },
            { "CopyPageSetup",           "PageSetup" },
            { "PagingSubtotals",         "Subtotal" },
            { "InsertPageBreakEveryRow", "PageBreakInsertExcel" },
            { "AddBorderToEachPage",     "OutsideBorders" },
        };

        public static string Build()
        {
            var sb = new StringBuilder();
            sb.Append("<customUI xmlns=\"").Append(Ns).Append("\" onLoad=\"OnRibbonLoad\"><ribbon><tabs>");

            var byTab = CommandRegistry.All
                .Select(c => c.Definition)
                .Where(d => d != null && !string.IsNullOrEmpty(d.Tab))
                .GroupBy(d => TabRemap.TryGetValue(d.Tab, out string mapped) ? mapped : d.Tab)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (string tab in OrderTabs(byTab.Keys))
            {
                AppendTab(sb, tab, byTab[tab]);
            }

            AppendSuiteTab(sb);

            sb.Append("</tabs></ribbon></customUI>");
            return sb.ToString();
        }

        private static IEnumerable<string> OrderTabs(IEnumerable<string> tabs)
        {
            var set = new List<string>(tabs);
            var ordered = new List<string>();
            foreach (string t in TabOrder)
                if (set.Contains(t)) { ordered.Add(t); set.Remove(t); }
            set.Sort(StringComparer.OrdinalIgnoreCase);
            ordered.AddRange(set);
            return ordered;
        }

        private static void AppendTab(StringBuilder sb, string tab, List<CommandDefinition> defs)
        {
            sb.Append("<tab id=\"").Append(SafeId("tab", tab)).Append("\" label=\"").Append(Esc(tab)).Append("\">");

            var groups = defs
                .Where(d => !string.IsNullOrEmpty(d.Group))
                .GroupBy(d => d.Group)
                .OrderBy(g => g.Min(d => d.Order));

            foreach (var group in groups)
            {
                sb.Append("<group id=\"").Append(SafeId("grp", tab + "_" + group.Key))
                  .Append("\" label=\"").Append(Esc(group.Key)).Append("\">");

                foreach (CommandDefinition d in group.OrderBy(x => x.Order).ThenBy(x => x.Label, StringComparer.OrdinalIgnoreCase))
                {
                    AppendButton(sb, d);
                }
                sb.Append("</group>");
            }
            sb.Append("</tab>");
        }

        private static void AppendSuiteTab(StringBuilder sb)
        {
            sb.Append("<tab id=\"tab_suite\" label=\"Suite\">");

            sb.Append("<group id=\"grp_history\" label=\"History\">");
            AppendDynamicButton(sb, RibbonController.SysUndoTag, true);
            sb.Append("</group>");

            sb.Append("<group id=\"grp_help\" label=\"Help\">");
            AppendStaticButton(sb, RibbonController.SysAboutTag, "About", false);
            AppendStaticButton(sb, RibbonController.SysOpenLogTag, "Open Log", false);
            sb.Append("</group>");

            sb.Append("</tab>");
        }

        private static void AppendButton(StringBuilder sb, CommandDefinition def)
        {
            // Resolve the icon: explicit ImageMso on the def wins; then check the
            // fallback table; finally fall back to the getImage/IconProvider path
            // (used by the 7 tools that have embedded custom PNGs).
            string mso = def.ImageMso;
            if (string.IsNullOrEmpty(mso) && !string.IsNullOrEmpty(def.ImageId))
                ImageIdToMso.TryGetValue(def.ImageId, out mso);

            sb.Append("<button id=\"").Append(SafeId("btn", def.Id)).Append("\"")
              .Append(" tag=\"").Append(Esc(def.Id)).Append("\"")
              .Append(" getLabel=\"GetLabel\" onAction=\"OnAction\"")
              .Append(" getScreentip=\"GetScreentip\" getSupertip=\"GetSupertip\"")
              .Append(" getEnabled=\"GetEnabled\"");

            if (!string.IsNullOrEmpty(mso))
                sb.Append(" imageMso=\"").Append(Esc(mso)).Append("\"");
            else
                sb.Append(" getImage=\"GetImage\"");

            sb.Append(" size=\"").Append(def.LargeButton ? "large" : "normal").Append("\"/>");
        }

        private static void AppendDynamicButton(StringBuilder sb, string tag, bool large)
        {
            sb.Append("<button id=\"").Append(SafeId("btn", tag)).Append("\"")
              .Append(" tag=\"").Append(Esc(tag)).Append("\"")
              .Append(" getLabel=\"GetLabel\" onAction=\"OnAction\" getEnabled=\"GetEnabled\"")
              .Append(" imageMso=\"Undo\"")
              .Append(" size=\"").Append(large ? "large" : "normal").Append("\"/>");
        }

        private static readonly Dictionary<string, string> SysButtonIcons =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "sys.about",   "Help" },
                { "sys.openlog", "FileOpen" },
            };

        private static void AppendStaticButton(StringBuilder sb, string tag, string label, bool large)
        {
            string mso;
            SysButtonIcons.TryGetValue(tag, out mso);
            sb.Append("<button id=\"").Append(SafeId("btn", tag)).Append("\"")
              .Append(" tag=\"").Append(Esc(tag)).Append("\"")
              .Append(" label=\"").Append(Esc(label)).Append("\" onAction=\"OnAction\"");
            if (!string.IsNullOrEmpty(mso))
                sb.Append(" imageMso=\"").Append(Esc(mso)).Append("\"");
            sb.Append(" size=\"").Append(large ? "large" : "normal").Append("\"/>");
        }

        /// <summary>Ribbon ids must be simple tokens; derive a stable safe id from arbitrary text.</summary>
        private static string SafeId(string prefix, string source)
        {
            var sb = new StringBuilder(prefix);
            sb.Append('_');
            foreach (char c in source)
            {
                sb.Append(char.IsLetterOrDigit(c) ? c : '_');
            }
            return sb.ToString();
        }

        private static string Esc(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");
        }
    }
}
