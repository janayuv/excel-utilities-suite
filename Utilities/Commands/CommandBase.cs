using System;
using System.Windows.Forms;
using utilities.Dialogs;
using utilities.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands
{
    /// <summary>
    /// Base for direct-action tools (those that operate immediately on the current scope
    /// without a dialog). Handles the license gate, target resolution and validation, then
    /// delegates the guarded mutation to <see cref="OperationRunner"/>. Concrete tools
    /// implement only <see cref="Run"/>.
    /// </summary>
    public abstract class CommandBase : IExcelCommand
    {
        public abstract CommandDefinition Definition { get; }

        public void Execute()
        {
            CommandDefinition def = Definition;

            if (!LicenseGate.Check(def)) return;

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

            OperationRunner.RunGuarded(def, target, Run);
        }

        /// <summary>The actual tool logic. Mutations are already guarded by OperationRunner.</summary>
        protected abstract void Run(CommandContext ctx);
    }

    /// <summary>
    /// Base for tools that open a dialog. The dialog gathers options and runs its own work
    /// through <see cref="OperationRunner"/> (via <c>DialogBase.RunOperation</c>), so undo
    /// and error handling stay identical to direct commands. Concrete tools implement only
    /// <see cref="CreateDialog"/>.
    /// </summary>
    public abstract class DialogCommandBase : IExcelCommand
    {
        public abstract CommandDefinition Definition { get; }

        public void Execute()
        {
            if (!LicenseGate.Check(Definition)) return;

            using (DialogBase dlg = CreateDialog())
            {
                if (dlg != null) dlg.ShowDialog();
            }
        }

        protected abstract DialogBase CreateDialog();
    }

    /// <summary>Shared license check + upgrade prompt used by both command bases.</summary>
    internal static class LicenseGate
    {
        public static bool Check(CommandDefinition def)
        {
            if (License.Current.IsFeatureAvailable(def.LicenseFeature)) return true;

            MessageBox.Show(
                "\"" + def.Label + "\" is not available in your current plan.\n\n" +
                License.Current.StatusText + "\n\nUpgrade to unlock this tool.",
                "Upgrade required",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            return false;
        }
    }

    /// <summary>
    /// Resolves the working range from a <see cref="CommandScope"/> and validates that a
    /// usable range is selected. Centralised so every tool reports "select a range" the
    /// same way (matching the existing add-in's behaviour).
    /// </summary>
    public static class RangeResolver
    {
        public static Excel.Range Resolve(CommandScope scope, out string error)
        {
            error = null;
            Excel.Application app = Globals.ThisAddIn.Application;

            if (app.ActiveWorkbook == null)
            {
                error = "Open a workbook first.";
                return null;
            }

            switch (scope)
            {
                case CommandScope.Worksheet:
                case CommandScope.Workbook:
                {
                    Excel.Worksheet ws = app.ActiveSheet as Excel.Worksheet;
                    if (ws == null) { error = "Select a worksheet first."; return null; }
                    Excel.Range used = ws.UsedRange;
                    if (used == null) { error = "The sheet has no data."; return null; }
                    return used;
                }

                default: // Selection
                {
                    Excel.Range sel = app.Selection as Excel.Range;
                    if (sel == null)
                    {
                        error = "Please select a range of cells first.";
                        return null;
                    }
                    return sel;
                }
            }
        }
    }
}
