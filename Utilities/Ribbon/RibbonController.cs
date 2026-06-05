using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Office.Core;
using utilities.Commands;
using utilities.Services;

namespace utilities.Ribbon
{
    /// <summary>
    /// Raw IRibbonExtensibility ribbon for the suite. The ribbon XML is generated from the
    /// command registry at runtime, so tabs/groups/buttons and their tooltips can never
    /// drift from the command definitions. All control callbacks dispatch by the control's
    /// tag (the command id).
    /// </summary>
    [System.Runtime.InteropServices.ComVisible(true)]
    public sealed class RibbonController : IRibbonExtensibility
    {
        // System (non-command) controls handled directly by the controller.
        private const string SysUndo = "sys.undo";
        private const string SysAbout = "sys.about";
        private const string SysOpenLog = "sys.openlog";
        private const string SysRepeat = "sys.repeat";
        private const string SysFindRun = "sys.findrun";

        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonId)
        {
            try
            {
                // GetCustomUI is called by Excel BEFORE ThisAddIn_Startup fires, so we must
                // ensure the registry is populated here rather than relying on Startup order.
                CommandRegistry.Initialize();
                return RibbonXmlBuilder.Build();
            }
            catch (Exception ex)
            {
                ErrorService.Log("ERROR", "Failed to build ribbon XML: " + ex);
                // Minimal fallback so Excel still loads the add-in.
                return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"/>";
            }
        }

        public void OnRibbonLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
        }

        /// <summary>Re-query dynamic callbacks (enabled state, undo label) — call on selection change.</summary>
        public void Invalidate()
        {
            try { if (_ribbon != null) _ribbon.Invalidate(); }
            catch { }
        }

        // ---- Shared callbacks (referenced from generated XML) ----

        public void OnAction(IRibbonControl control)
        {
            string id = TagOf(control);
            switch (id)
            {
                case SysUndo:
                    UndoService.UndoLast();
                    Invalidate();
                    return;
                case SysAbout:
                    ShowAbout();
                    return;
                case SysOpenLog:
                    OpenLog();
                    return;
                case SysRepeat:
                    RepeatService.Replay();
                    Invalidate();
                    return;
                case SysFindRun:
                    ShowFindRun();
                    return;
            }

            IExcelCommand cmd = CommandRegistry.Get(id);
            if (cmd == null)
            {
                ErrorService.Log("WARN", "No command for ribbon action '" + id + "'.");
                return;
            }
            cmd.Execute();
            Invalidate();
        }

        public string GetLabel(IRibbonControl control)
        {
            string id = TagOf(control);
            if (id == SysUndo)
            {
                string next = UndoService.NextUndoLabel;
                return next != null ? "Undo " + next : "Undo Last Action";
            }
            if (id == SysRepeat)
            {
                return RepeatService.CanRepeat
                    ? "Repeat: " + RepeatService.LastLabel
                    : "Repeat Last Tool";
            }
            CommandDefinition def = CommandRegistry.GetDefinition(id);
            return def != null ? def.Label : id;
        }

        public string GetScreentip(IRibbonControl control)
        {
            CommandDefinition def = CommandRegistry.GetDefinition(TagOf(control));
            return def != null ? (def.Screentip ?? def.Label) : null;
        }

        public string GetSupertip(IRibbonControl control)
        {
            CommandDefinition def = CommandRegistry.GetDefinition(TagOf(control));
            return def != null ? def.Supertip : null;
        }

        public stdole.IPictureDisp GetImage(IRibbonControl control)
        {
            string id = TagOf(control);
            CommandDefinition def = CommandRegistry.GetDefinition(id);
            string imageId = def != null ? def.ImageId : null;
            return IconProvider.GetPicture(imageId);
        }

        public bool GetEnabled(IRibbonControl control)
        {
            string id = TagOf(control);
            if (id == SysUndo) return UndoService.CanUndo;
            if (id == SysRepeat) return RepeatService.CanRepeat;
            return true; // license-locked tools still show; they prompt to upgrade on click
        }

        // ---- System control handlers ----

        public bool GetLicenseVisible(IRibbonControl control)
        {
            return License.Current.State != LicenseState.Licensed;
        }

        private void ShowFindRun()
        {
            string id = null;
            using (var dlg = new utilities.Dialogs.FindRunDialog())
            {
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    id = dlg.SelectedCommandId;
            }

            if (!string.IsNullOrEmpty(id))
            {
                IExcelCommand cmd = CommandRegistry.Get(id);
                if (cmd != null) cmd.Execute(); // flows through RunGuarded -> updates Repeat
            }
            Invalidate();
        }

        private void ShowAbout()
        {
            var svc = License.Current;
            var sb = new StringBuilder();
            sb.AppendLine("Excel Utilities Suite");
            sb.AppendLine(CommandRegistry.Count + " tools installed.");
            sb.AppendLine();
            sb.AppendLine("─── License ───────────────────────");
            sb.AppendLine("Status:     " + svc.StatusText);
            if (svc.State == LicenseState.Licensed || svc.State == LicenseState.Offline)
            {
                sb.AppendLine("Machine ID: " + RealLicenseService.MachineId());
            }
            System.Windows.Forms.MessageBox.Show(
                sb.ToString(),
                "About Excel Utilities Suite",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private void OpenLog()
        {
            try
            {
                if (System.IO.File.Exists(ErrorService.LogFilePath))
                    Process.Start(ErrorService.LogFilePath);
                else
                    System.Windows.Forms.MessageBox.Show("No log entries yet.", "Log");
            }
            catch (Exception ex)
            {
                ErrorService.Log("WARN", "Could not open log: " + ex.Message);
            }
        }

        private static string TagOf(IRibbonControl control)
        {
            return control != null ? (control.Tag ?? control.Id) : null;
        }

        // System control tags used by the XML builder.
        internal static string SysUndoTag { get { return SysUndo; } }
        internal static string SysAboutTag { get { return SysAbout; } }
        internal static string SysOpenLogTag { get { return SysOpenLog; } }
        internal static string SysRepeatTag { get { return SysRepeat; } }
        internal static string SysFindRunTag { get { return SysFindRun; } }
    }
}
