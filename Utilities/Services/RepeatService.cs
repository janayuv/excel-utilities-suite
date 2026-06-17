using System;
using System.Collections.Generic;
using System.Windows.Forms;
using utilities.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Services
{
    /// <summary>
    /// Records the most recent repeatable tool invocation and replays it silently with the
    /// same settings against the current selection ("Repeat Last Tool" / Excel-F4 behaviour).
    ///
    /// Capture point is <see cref="OperationRunner.RunGuarded"/>: every direct and dialog tool
    /// funnels its work closure through there, and that closure already has the user's chosen
    /// settings baked in and operates on ctx.Target — so replaying it against a freshly resolved
    /// target re-applies the same operation with no per-tool code.
    /// </summary>
    public static class RepeatService
    {
        private static CommandDefinition _lastDef;
        private static Action<CommandContext> _lastWork;

        // Tool groups whose operations must not silently re-fire (file dialogs, printing).
        // An excluded tool leaves the previous "last" pointer unchanged.
        private static readonly HashSet<string> NonRepeatableGroups =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Printing", "Export", "Import"
            };

        // Specific multi-range / multi-file commands that are unsafe to replay silently.
        private static readonly HashSet<string> NonRepeatableIds =
            new HashSet<string>(StringComparer.Ordinal)
            {
                "Range.SwapRanges",
                "Range.CompareRanges",
                "Sheet.MergeWorkbooks",
                "Sheet.SplitByColumn"
            };

        /// <summary>True when a replayable action is stored.</summary>
        public static bool CanRepeat { get { return _lastDef != null && _lastWork != null; } }

        /// <summary>Label of the last repeatable tool, or null.</summary>
        public static string LastLabel { get { return _lastDef != null ? _lastDef.Label : null; } }

        /// <summary>Id of the last repeatable tool, or null.</summary>
        public static string LastId { get { return _lastDef != null ? _lastDef.Id : null; } }

        /// <summary>
        /// Record a successful tool invocation as the new "last action", unless the tool is
        /// non-repeatable. Called from OperationRunner.RunGuarded on success.
        /// </summary>
        public static void Record(CommandDefinition def, Action<CommandContext> work)
        {
            if (def == null || work == null) return;
            if (!IsRepeatable(def)) return; // leave the previous last-action intact
            _lastDef = def;
            _lastWork = work;
        }

        /// <summary>
        /// Forget the stored last action. Called when a workbook closes so a captured work
        /// closure cannot replay against a different workbook (or hold a stale COM reference).
        /// </summary>
        public static void Clear()
        {
            _lastDef = null;
            _lastWork = null;
        }

        /// <summary>Pure predicate: may this tool be silently repeated?</summary>
        public static bool IsRepeatable(CommandDefinition def)
        {
            if (def == null) return false;
            if (!string.IsNullOrEmpty(def.Group) && NonRepeatableGroups.Contains(def.Group)) return false;
            if (!string.IsNullOrEmpty(def.Id) && NonRepeatableIds.Contains(def.Id)) return false;
            return true;
        }

        /// <summary>
        /// Re-run the last recorded tool with its captured settings against the current
        /// selection. Shows the standard "select a range" warning when a target is required
        /// but missing.
        /// </summary>
        public static void Replay()
        {
            if (!CanRepeat) return;

            CommandDefinition def = _lastDef;
            Action<CommandContext> work = _lastWork;

            Excel.Range target = null;
            if (def.RequiresSelection)
            {
                string error;
                target = RangeResolver.Resolve(def.Scope, out error);
                if (target == null)
                {
                    MessageBox.Show(error, def.Label, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            OperationRunner.RunGuarded(def, target, work);
        }
    }
}
