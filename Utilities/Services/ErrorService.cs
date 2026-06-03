using System;
using System.IO;
using System.Windows.Forms;
using utilities.Commands;

namespace utilities.Services
{
    /// <summary>
    /// Single place that turns an exception into a friendly dialog and a rolling log
    /// entry under %APPDATA%. Every command funnels failures here via CommandBase so the
    /// user experience is uniform.
    /// </summary>
    public static class ErrorService
    {
        private const string AppFolderName = "ExcelUtilitiesSuite";
        private const string LogFileName = "suite.log";

        public static string LogDirectory
        {
            get
            {
                string root = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                return Path.Combine(root, AppFolderName);
            }
        }

        public static string LogFilePath
        {
            get { return Path.Combine(LogDirectory, LogFileName); }
        }

        /// <summary>Show a friendly error for a failed command and log the detail.</summary>
        public static void Handle(Exception ex, CommandDefinition def)
        {
            string toolName = def != null ? def.Label : "Operation";
            Log("ERROR", (def != null ? def.Id : "?") + ": " + ex);

            MessageBox.Show(
                toolName + " could not be completed.\n\n" + ex.Message +
                "\n\nDetails were written to the log (Help → Open Log).",
                toolName,
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        /// <summary>Append a timestamped line to the rolling log (best-effort, never throws).</summary>
        public static void Log(string level, string message)
        {
            try
            {
                Directory.CreateDirectory(LogDirectory);

                // Roll the log when it grows past ~1 MB.
                try
                {
                    var fi = new FileInfo(LogFilePath);
                    if (fi.Exists && fi.Length > 1000000)
                    {
                        string archived = LogFilePath + ".1";
                        if (File.Exists(archived)) File.Delete(archived);
                        File.Move(LogFilePath, archived);
                    }
                }
                catch { }

                File.AppendAllText(LogFilePath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + level + "] " + message + Environment.NewLine);
            }
            catch
            {
                // Logging must never break the add-in.
            }
        }
    }
}
