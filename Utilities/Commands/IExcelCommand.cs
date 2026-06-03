namespace utilities.Commands
{
    /// <summary>
    /// Contract for every tool in the suite. A command exposes its metadata through a
    /// single <see cref="Definition"/> (the source of truth) and implements only the
    /// work in <see cref="Run"/>; all guard/perf/undo/license plumbing lives in
    /// <c>CommandBase</c>.
    /// </summary>
    public interface IExcelCommand
    {
        CommandDefinition Definition { get; }

        /// <summary>
        /// Entry point invoked by the ribbon. Implemented by <c>CommandBase</c> (direct
        /// tools) or <c>DialogCommandBase</c> (dialog tools); concrete commands supply the
        /// actual work via <c>Run</c> or <c>CreateDialog</c> respectively.
        /// </summary>
        void Execute();
    }
}
