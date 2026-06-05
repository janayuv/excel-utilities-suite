using System;
using utilities.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands
{
    /// <summary>
    /// The one place the suite's cross-cutting execution concerns live: the performance
    /// toggle (ScreenUpdating/Calculation/EnableEvents), status-bar progress, undo capture
    /// driven by the declared <see cref="UndoMode"/>, friendly error handling, and the
    /// best-effort rollback on failure.
    ///
    /// Both <see cref="CommandBase"/> (direct tools) and dialogs (via <c>DialogBase</c>)
    /// run their mutations through here so behaviour is identical everywhere.
    /// </summary>
    public static class OperationRunner
    {
        /// <summary>
        /// Execute <paramref name="work"/> against <paramref name="target"/> with full
        /// guarding. Returns true on success.
        /// </summary>
        public static bool RunGuarded(CommandDefinition def, Excel.Range target, Action<CommandContext> work)
        {
            if (def == null || work == null) return false;

            Excel.Application app = Globals.ThisAddIn.Application;

            var progress = new StatusBarProgressReporter(app, def.Label);
            UndoTransaction tx = UndoService.Begin(def.Label, def.UndoMode, target);

            // Snapshot Excel state so we always restore it, even on error.
            bool prevScreen = true;
            bool prevEvents = true;
            Excel.XlCalculation prevCalc = Excel.XlCalculation.xlCalculationAutomatic;
            bool calcCaptured = false;

            try
            {
                prevScreen = app.ScreenUpdating;
                prevEvents = app.EnableEvents;
                try { prevCalc = app.Calculation; calcCaptured = true; } catch { /* no workbook */ }

                app.ScreenUpdating = false;
                app.EnableEvents = false;
                if (calcCaptured) { try { app.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { } }

                var ctx = new CommandContext(app, target, def, progress, tx);
                work(ctx);

                UndoService.Push(tx);
                // Capture this invocation for "Repeat Last Tool". Safe against chained
                // replays: RepeatService.Replay re-invokes RunGuarded with the ORIGINAL
                // captured work closure, so repeating never recurses into Replay itself.
                RepeatService.Record(def, work);
                return true;
            }
            catch (Exception ex)
            {
                try { tx.Restore(); } catch { /* rollback is best-effort */ }
                ErrorService.Handle(ex, def);
                return false;
            }
            finally
            {
                try { app.ScreenUpdating = prevScreen; } catch { }
                try { app.EnableEvents = prevEvents; } catch { }
                if (calcCaptured) { try { app.Calculation = prevCalc; } catch { } }
                progress.Done();
            }
        }
    }
}
