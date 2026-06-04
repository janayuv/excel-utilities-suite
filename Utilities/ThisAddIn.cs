using System;
using Microsoft.Office.Core;
using utilities.Commands;
using utilities.Ribbon;
using utilities.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities
{
    public partial class ThisAddIn
    {
        private RibbonController _ribbon;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                License.Current = RealLicenseService.Load();
                CommandRegistry.Initialize();

                // Keep the Undo button's enabled-state and label in sync with the selection.
                this.Application.SheetSelectionChange += OnSheetSelectionChange;

                // Release the cached sheet COM reference before the workbook is gone.
                this.Application.WorkbookBeforeClose += OnWorkbookBeforeClose;

                // Intercept Ctrl+Z: consume it when the add-in stack has something,
                // otherwise the key falls through to Excel's native undo.
                KeyboardHook.Install(() =>
                {
                    UndoService.UndoLast();
                    if (_ribbon != null) _ribbon.Invalidate();
                });

                ErrorService.Log("INFO", "Add-in started. License: " + License.Current.StatusText);
            }
            catch (Exception ex)
            {
                ErrorService.Log("ERROR", "Startup failed: " + ex);
            }
        }

        private void OnWorkbookBeforeClose(Excel.Workbook wb, ref bool cancel)
        {
            GridFocusService.ClearLastSheet();
        }

        private void OnSheetSelectionChange(object sh, Excel.Range target)
        {
            GridFocusService.OnSelectionChange(this.Application);
            if (_ribbon != null) _ribbon.Invalidate();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            KeyboardHook.Uninstall();
            try { GridFocusService.ClearAll(this.Application); } catch { }
            try { this.Application.SheetSelectionChange -= OnSheetSelectionChange; } catch { }
            try { this.Application.WorkbookBeforeClose -= OnWorkbookBeforeClose; } catch { }
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new RibbonController();
            return _ribbon;
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
