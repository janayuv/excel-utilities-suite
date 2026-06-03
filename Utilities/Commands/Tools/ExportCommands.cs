using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    /// <summary>
    /// Export the current selection (if more than one cell) or the active sheet to a PDF file.
    /// </summary>
    [ExcelCommand]
    public sealed class ExportToPdfCommand : CommandBase
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Export.ToPdf",
            Label = "Export to PDF",
            Screentip = "Export to PDF",
            Supertip = "Save the selected range, or the active worksheet, as a PDF document.",
            ImageId = "ExportToPDF",
            Tab = "Export / Import",
            Group = "Export",
            Order = 10,
            Scope = CommandScope.Selection,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override void Run(CommandContext ctx)
        {
            Excel.Application app = ctx.App;
            Excel.Range selection = app.Selection as Excel.Range;
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;

            bool exportSelection = selection != null && selection.Cells.Count > 1;

            string suggested = (sheet != null ? sheet.Name : "Export");
            using (var dlg = new SaveFileDialog
            {
                Title = "Export to PDF",
                Filter = "PDF document (*.pdf)|*.pdf",
                FileName = SanitizeFileName(suggested) + ".pdf"
            })
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                object source = exportSelection ? (object)selection : sheet;
                ((dynamic)source).ExportAsFixedFormat(
                    Excel.XlFixedFormatType.xlTypePDF,
                    dlg.FileName,
                    Excel.XlFixedFormatQuality.xlQualityStandard,
                    true,  // include doc properties
                    false, // ignore print areas
                    Type.Missing, Type.Missing, false, Type.Missing);

                MessageBox.Show("Exported to:\n" + dlg.FileName, Definition.Label,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');
            return name;
        }
    }
}
