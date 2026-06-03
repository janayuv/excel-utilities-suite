using System;
using utilities.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands
{
    /// <summary>
    /// Everything a command's <c>Run</c> needs: the Excel application, the resolved
    /// target range, the live definition, a progress reporter and the undo recorder.
    /// Created by <see cref="CommandBase"/> so individual commands never repeat the
    /// selection/guard/perf plumbing.
    /// </summary>
    public sealed class CommandContext
    {
        public Excel.Application App { get; }

        /// <summary>
        /// The range the command should act on, already resolved from the declared
        /// <see cref="CommandScope"/>. Null when the command declares
        /// <c>RequiresSelection = false</c>.
        /// </summary>
        public Excel.Range Target { get; }

        public CommandDefinition Definition { get; }

        public IProgressReporter Progress { get; }

        /// <summary>
        /// For <see cref="UndoMode.PartialSnapshot"/> commands: call before mutating a cell
        /// so the previous value can be restored. No-op for other undo modes.
        /// </summary>
        public IUndoRecorder Undo { get; }

        public CommandContext(Excel.Application app, Excel.Range target, CommandDefinition definition,
            IProgressReporter progress, IUndoRecorder undo)
        {
            App = app;
            Target = target;
            Definition = definition;
            Progress = progress;
            Undo = undo;
        }
    }
}
