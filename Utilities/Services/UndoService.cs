using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands
{
    // Note: kept in the utilities.Commands namespace because CommandContext/CommandBase
    // depend on IUndoRecorder and it avoids a circular using between the two namespaces.

    /// <summary>
    /// Passed to a command so it can record a cell's prior value immediately before
    /// mutating it. Only <see cref="UndoMode.PartialSnapshot"/> commands use this; for
    /// other modes the calls are harmless no-ops.
    /// </summary>
    public interface IUndoRecorder
    {
        void RecordCell(Excel.Range cell);
    }
}

namespace utilities.Services
{
    using utilities.Commands;

    /// <summary>
    /// Captures the state needed to reverse a single command and exposes an undo stack.
    /// Restore is driven by the suite's own "Undo Last Action" button rather than by
    /// hijacking Excel's native Ctrl+Z, which is unreliable for managed add-ins.
    /// </summary>
    public static class UndoService
    {
        /// <summary>
        /// Above this many cells a FullSnapshot is refused (the command runs without undo
        /// after warning). Keeps memory realistic on 50k+ ranges.
        /// </summary>
        public const int FullSnapshotCellLimit = 200000;

        private static readonly Stack<UndoTransaction> _stack = new Stack<UndoTransaction>();

        public static bool CanUndo { get { return _stack.Count > 0; } }

        public static string NextUndoLabel
        {
            get { return _stack.Count > 0 ? _stack.Peek().Label : null; }
        }

        /// <summary>
        /// Begin capturing for a command. Returns a transaction that is also the
        /// <see cref="IUndoRecorder"/> handed to the command. Up-front modes snapshot here;
        /// PartialSnapshot captures lazily as the command records cells.
        /// </summary>
        public static UndoTransaction Begin(string label, UndoMode mode, Excel.Range target)
        {
            var tx = new UndoTransaction(label, mode, target);
            tx.CaptureUpFront();
            return tx;
        }

        /// <summary>Push a completed transaction onto the undo stack (cap depth at 20).</summary>
        public static void Push(UndoTransaction tx)
        {
            if (tx == null || tx.Mode == UndoMode.None || tx.IsEmpty) return;
            _stack.Push(tx);
            while (_stack.Count > 20)
            {
                // Drop the oldest by rebuilding (Stack has no remove-last).
                var keep = new UndoTransaction[20];
                for (int i = 19; i >= 0; i--) keep[i] = _stack.Pop();
                _stack.Clear();
                for (int i = 0; i < 20; i++) _stack.Push(keep[i]);
                break;
            }
        }

        /// <summary>Restore the most recent transaction.</summary>
        public static void UndoLast()
        {
            if (_stack.Count == 0) return;
            UndoTransaction tx = _stack.Pop();
            tx.Restore();
        }

        public static void Clear() { _stack.Clear(); }
    }

    /// <summary>One reversible unit of work captured for a command.</summary>
    public sealed class UndoTransaction : IUndoRecorder
    {
        private readonly Excel.Range _target;
        private readonly Dictionary<string, object> _cellValues = new Dictionary<string, object>();
        private object[,] _blockValues;     // for FullSnapshot/FormulaOnly: rectangular capture
        private string _worksheetName;
        private string _blockAddress;

        public string Label { get; }
        public UndoMode Mode { get; }
        public bool Downgraded { get; private set; }

        public bool IsEmpty
        {
            get { return _blockValues == null && _cellValues.Count == 0; }
        }

        public UndoTransaction(string label, UndoMode mode, Excel.Range target)
        {
            Label = label;
            Mode = mode;
            _target = target;
        }

        /// <summary>Capture rectangular state for the up-front modes.</summary>
        public void CaptureUpFront()
        {
            if (_target == null) return;
            if (Mode != UndoMode.FullSnapshot && Mode != UndoMode.FormulaOnly) return;

            int count = 0;
            try { count = _target.Cells.Count; } catch { count = int.MaxValue; }

            if (Mode == UndoMode.FullSnapshot && count > UndoService.FullSnapshotCellLimit)
            {
                Downgraded = true; // too large; skip capture, command runs without undo
                return;
            }

            try
            {
                _worksheetName = ((Excel.Worksheet)_target.Worksheet).Name;
                _blockAddress = _target.Address;
                if (Mode == UndoMode.FormulaOnly)
                {
                    _blockValues = _target.Formula as object[,];
                    if (_blockValues == null) // single cell returns a scalar
                    {
                        _cellValues[_target.Address] = _target.Formula;
                    }
                }
                else
                {
                    _blockValues = _target.Value2 as object[,];
                    if (_blockValues == null)
                    {
                        _cellValues[_target.Address] = _target.Value2;
                    }
                }
            }
            catch
            {
                _blockValues = null;
            }
        }

        /// <summary>PartialSnapshot: remember a cell's value the first time it is touched.</summary>
        public void RecordCell(Excel.Range cell)
        {
            if (Mode != UndoMode.PartialSnapshot || cell == null) return;
            try
            {
                string key = cell.Address;
                if (!_cellValues.ContainsKey(key))
                {
                    _cellValues[key] = cell.Value2;
                }
            }
            catch { }
        }

        /// <summary>Put the captured state back.</summary>
        public void Restore()
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            bool prevEvents = true, prevScreen = true;
            try
            {
                prevEvents = app.EnableEvents; prevScreen = app.ScreenUpdating;
                app.EnableEvents = false; app.ScreenUpdating = false;

                Excel.Worksheet sheet = _worksheetName != null
                    ? FindSheet(app, _worksheetName)
                    : app.ActiveSheet as Excel.Worksheet;

                if (_blockValues != null && sheet != null && _blockAddress != null)
                {
                    Excel.Range block = sheet.Range[_blockAddress];
                    if (Mode == UndoMode.FormulaOnly) block.Formula = _blockValues;
                    else block.Value2 = _blockValues;
                }

                foreach (var kv in _cellValues)
                {
                    Excel.Range cell = sheet != null ? sheet.Range[kv.Key] : null;
                    if (cell == null) continue;
                    if (Mode == UndoMode.FormulaOnly) cell.Formula = kv.Value;
                    else cell.Value2 = kv.Value;
                }
            }
            finally
            {
                try { app.EnableEvents = prevEvents; app.ScreenUpdating = prevScreen; } catch { }
            }
        }

        private static Excel.Worksheet FindSheet(Excel.Application app, string name)
        {
            try
            {
                foreach (Excel.Worksheet ws in app.ActiveWorkbook.Worksheets)
                {
                    if (ws.Name == name) return ws;
                }
            }
            catch { }
            return app.ActiveSheet as Excel.Worksheet;
        }
    }
}
