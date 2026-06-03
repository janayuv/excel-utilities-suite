using System;
using System.Windows.Forms;

namespace utilities.Commands.Tools
{
    /// <summary>
    /// Thin command wrappers around the original WinForms dialogs (SequenceForm,
    /// ChangeCaseForm, TrimTextForm). They reuse the existing forms unchanged for behaviour
    /// parity; the forms perform their own work, so these go through the license gate and
    /// open the dialog directly rather than via OperationRunner.
    /// </summary>
    internal abstract class LegacyDialogCommand : IExcelCommand
    {
        public abstract CommandDefinition Definition { get; }

        public void Execute()
        {
            if (!LicenseGate.Check(Definition)) return;
            using (Form form = CreateForm())
            {
                if (form != null) form.ShowDialog();
            }
        }

        protected abstract Form CreateForm();
    }

    [ExcelCommand]
    internal sealed class InsertSequenceCommand : LegacyDialogCommand
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Insert.SequenceNumbers",
            Label = "Insert Sequence",
            Screentip = "Insert Sequence Numbers",
            Supertip = "Fill the selection with a custom sequence of numbers (start value, step, prefix/suffix).",
            ImageId = "InsertSequence",
            Tab = "Insert",
            Group = "Insert",
            Order = 10,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override Form CreateForm() { return new SequenceForm(); }
    }

    [ExcelCommand]
    internal sealed class ChangeCaseCommand : LegacyDialogCommand
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.ChangeCase",
            Label = "Change Case",
            Screentip = "Change Case",
            Supertip = "Convert the text in the selection to UPPER, lower, Proper or Sentence case.",
            ImageId = "ChangeCase",
            Tab = "Editing",
            Group = "Text",
            Order = 10,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override Form CreateForm() { return new ChangeCaseForm(); }
    }

    [ExcelCommand]
    internal sealed class TrimTextCommand : LegacyDialogCommand
    {
        public override CommandDefinition Definition { get; } = new CommandDefinition
        {
            Id = "Text.TrimText",
            Label = "Trim / Clean Text",
            Screentip = "Trim and Clean Text",
            Supertip = "Remove leading, trailing and excessive spaces, or delete a set number of leading/ending characters.",
            ImageId = "TrimText",
            Tab = "Editing",
            Group = "Text",
            Order = 11,
            RequiresSelection = false,
            UndoMode = UndoMode.None
        };

        protected override Form CreateForm() { return new TrimTextForm(); }
    }
}
