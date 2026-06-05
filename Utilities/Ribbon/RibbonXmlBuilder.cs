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

        // imageMso for each ribbon group — used when Office collapses the group into a single button.
        private static readonly Dictionary<string, string> GroupImageMso = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "View",              "ViewGridlines" },
            { "Format & Convert",  "ClearFormats" },
            { "Text",              "Bold" },
            { "Select",            "GoToSpecial" },
            { "Navigate",          "GoTo" },
            { "Insert",            "InsertFunction" },
            { "Range & Cells",     "ColumnWidthAutoFit" },
            { "Statistics",        "AutoSum" },
            { "Formulas",          "FunctionWizard" },
            { "Duplicates",        "RemoveDuplicates" },
            { "Merge",             "MergeCenter" },
            { "Combine",           "GroupRows" },
            { "Split",             "TextToColumns" },
            { "Export",            "SaveAs" },
            { "Import",            "FileOpen" },
            { "Sheets",            "InsertWorksheet" },
            { "Workbook",          "OpenDocumentExcelWorkbook" },
            { "Printing",          "FilePrint" },
            { "Quality",           "DataValidation" },
        };

        // All tool buttons now use getImage="GetImage" → IconProvider → custom embedded PNGs.
        // imageMso is only used for system controls (Undo, About, Open Log) and
        // collapsed-group icons (GroupImageMso above).

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
                string grpMso;
                GroupImageMso.TryGetValue(group.Key, out grpMso);
                sb.Append("<group id=\"").Append(SafeId("grp", tab + "_" + group.Key))
                  .Append("\" label=\"").Append(Esc(group.Key)).Append("\"");
                if (!string.IsNullOrEmpty(grpMso))
                    sb.Append(" imageMso=\"").Append(Esc(grpMso)).Append("\"");
                sb.Append(">");

                // Separate standalone buttons from commands that live inside a <menu> dropdown.
                var standalone = group.Where(d => string.IsNullOrEmpty(d.MenuParent)).ToList();
                var menuGroups = group.Where(d => !string.IsNullOrEmpty(d.MenuParent))
                                      .GroupBy(d => d.MenuParent).ToList();

                // Build ordered (effectiveOrder, renderAction) list so menus and buttons share one ordering pass.
                var items = new List<System.Tuple<int, Action>>();
                foreach (var d in standalone)
                {
                    var captured = d;
                    items.Add(System.Tuple.Create(d.Order, (Action)(() => AppendButton(sb, captured))));
                }
                foreach (var mg in menuGroups)
                {
                    var children = mg.OrderBy(d => d.Order).ThenBy(d => d.Label, StringComparer.OrdinalIgnoreCase).ToList();
                    int menuOrder = children.Any(d => d.MenuParentOrder > 0)
                        ? children.Where(d => d.MenuParentOrder > 0).Min(d => d.MenuParentOrder)
                        : children.Min(d => d.Order);
                    string parentLabel = mg.Key;
                    string parentMso   = children.Select(d => d.MenuParentImageMso).FirstOrDefault(s => !string.IsNullOrEmpty(s));
                    var capturedChildren = children;
                    items.Add(System.Tuple.Create(menuOrder,
                        (Action)(() => AppendMenuButton(sb, tab + "_" + group.Key + "_" + parentLabel, parentLabel, parentMso, capturedChildren))));
                }

                foreach (var item in items.OrderBy(x => x.Item1))
                    item.Item2();

                sb.Append("</group>");
            }
            sb.Append("</tab>");
        }

        private static void AppendMenuButton(StringBuilder sb, string idBase, string label, string imageMso, List<CommandDefinition> children)
        {
            sb.Append("<menu id=\"").Append(SafeId("menu", idBase)).Append("\"");
            sb.Append(" label=\"").Append(Esc(label)).Append("\"");
            if (!string.IsNullOrEmpty(imageMso))
                sb.Append(" imageMso=\"").Append(Esc(imageMso)).Append("\"");
            sb.Append(" size=\"large\">");
            foreach (var child in children)
            {
                sb.Append("<button id=\"").Append(SafeId("mitem", child.Id)).Append("\"")
                  .Append(" tag=\"").Append(Esc(child.Id)).Append("\"")
                  .Append(" getLabel=\"GetLabel\" onAction=\"OnAction\"")
                  .Append(" getScreentip=\"GetScreentip\" getSupertip=\"GetSupertip\"")
                  .Append(" getEnabled=\"GetEnabled\"");
                if (!string.IsNullOrEmpty(child.ImageMso))
                    sb.Append(" imageMso=\"").Append(Esc(child.ImageMso)).Append("\"");
                else if (!string.IsNullOrEmpty(child.ImageId))
                    sb.Append(" getImage=\"GetImage\"");
                sb.Append("/>");
            }
            sb.Append("</menu>");
        }

        private static void AppendSuiteTab(StringBuilder sb)
        {
            sb.Append("<tab id=\"tab_suite\" label=\"Suite\">");

            sb.Append("<group id=\"grp_quick\" label=\"Quick\">");
            AppendDynamicButton(sb, RibbonController.SysRepeatTag, "Repeat", true);
            sb.Append("</group>");

            // License group — hidden automatically once a valid key is stored.
            sb.Append("<group id=\"grp_license\" label=\"License\" getVisible=\"GetLicenseVisible\">");
            sb.Append("<button id=\"btn_Suite_ActivateLicense\" tag=\"Suite.ActivateLicense\"");
            sb.Append(" label=\"Activate License\" onAction=\"OnAction\" imageMso=\"Permissions\" size=\"normal\"/>");
            sb.Append("<button id=\"btn_Suite_DeactivateLicense\" tag=\"Suite.DeactivateLicense\"");
            sb.Append(" label=\"Deactivate\" onAction=\"OnAction\" imageMso=\"Delete\" size=\"normal\"/>");
            sb.Append("</group>");

            sb.Append("<group id=\"grp_history\" label=\"History\">");
            AppendDynamicButton(sb, RibbonController.SysUndoTag, "Undo", true);
            sb.Append("</group>");

            sb.Append("<group id=\"grp_help\" label=\"Help\">");
            AppendStaticButton(sb, RibbonController.SysAboutTag, "About", false);
            AppendStaticButton(sb, RibbonController.SysOpenLogTag, "Open Log", false);
            sb.Append("</group>");

            sb.Append("</tab>");
        }

        private static void AppendButton(StringBuilder sb, CommandDefinition def)
        {
            // All tool buttons use custom embedded PNGs via getImage="GetImage".
            // IconProvider selects light or dark icon automatically based on the Office theme.
            sb.Append("<button id=\"").Append(SafeId("btn", def.Id)).Append("\"")
              .Append(" tag=\"").Append(Esc(def.Id)).Append("\"")
              .Append(" getLabel=\"GetLabel\" onAction=\"OnAction\"")
              .Append(" getScreentip=\"GetScreentip\" getSupertip=\"GetSupertip\"")
              .Append(" getEnabled=\"GetEnabled\"")
              .Append(" getImage=\"GetImage\"")
              .Append(" size=\"").Append(def.LargeButton ? "large" : "normal").Append("\"/>");
        }

        private static void AppendDynamicButton(StringBuilder sb, string tag, string imageMso, bool large)
        {
            sb.Append("<button id=\"").Append(SafeId("btn", tag)).Append("\"")
              .Append(" tag=\"").Append(Esc(tag)).Append("\"")
              .Append(" getLabel=\"GetLabel\" onAction=\"OnAction\" getEnabled=\"GetEnabled\"")
              .Append(" imageMso=\"").Append(Esc(imageMso)).Append("\"")
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
