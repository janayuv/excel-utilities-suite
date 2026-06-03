using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Services
{
    /// <summary>
    /// Lightweight progress channel passed into every command. Implementations report to
    /// the Excel status bar today; a modal cancelable dialog can be swapped in for very
    /// long operations without changing command code.
    /// </summary>
    public interface IProgressReporter
    {
        /// <summary>Report progress as a fraction 0..1 with an optional message.</summary>
        void Report(double fraction, string message = null);

        /// <summary>True when the user requested cancellation.</summary>
        bool IsCancellationRequested { get; }

        /// <summary>Clear any progress UI.</summary>
        void Done();
    }

    /// <summary>Reports progress through <c>Application.StatusBar</c>, matching the existing add-in idiom.</summary>
    public sealed class StatusBarProgressReporter : IProgressReporter
    {
        private readonly Excel.Application _app;
        private readonly string _label;
        private double _lastShown = -1;

        public StatusBarProgressReporter(Excel.Application app, string label)
        {
            _app = app;
            _label = label ?? "Working";
        }

        public bool IsCancellationRequested
        {
            // Status-bar reporting has no cancel affordance; the modal reporter (future) will.
            get { return false; }
        }

        public void Report(double fraction, string message = null)
        {
            if (fraction < 0) fraction = 0;
            if (fraction > 1) fraction = 1;

            // Throttle status-bar writes to whole-percent changes to avoid COM chatter.
            double pct = Math.Round(fraction * 100);
            if (pct == _lastShown) return;
            _lastShown = pct;

            try
            {
                _app.StatusBar = (message ?? _label) + ": " + pct + "%";
            }
            catch
            {
                // Status bar is best-effort; never let it break a command.
            }
        }

        public void Done()
        {
            try { _app.StatusBar = false; } catch { }
        }
    }
}
