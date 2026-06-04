using System;

namespace utilities.Commands
{
    /// <summary>
    /// Where a command operates by default. The user can override scope in dialogs
    /// that expose a scope selector.
    /// </summary>
    public enum CommandScope
    {
        Selection,
        Worksheet,
        Workbook
    }

    /// <summary>
    /// How much state a command captures so it can be undone. Declared per command
    /// (see <see cref="CommandDefinition.UndoMode"/>) instead of assuming a full
    /// snapshot for every tool, which is unrealistic on large ranges.
    /// </summary>
    public enum UndoMode
    {
        /// <summary>No undo captured (read-only / select / navigate / export tools).</summary>
        None,

        /// <summary>Capture Value2 + key formats of the whole target range (small ranges only).</summary>
        FullSnapshot,

        /// <summary>Capture only the cells actually mutated, recorded during Run.</summary>
        PartialSnapshot,

        /// <summary>Capture/restore formulas only (for formula-rewriting tools).</summary>
        FormulaOnly
    }

    /// <summary>
    /// Marker attribute used by <c>CommandRegistry</c> to discover concrete commands
    /// via reflection. Only types carrying this attribute are registered, which keeps
    /// base/helper classes out of the registry.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public sealed class ExcelCommandAttribute : Attribute
    {
    }

    /// <summary>
    /// The single source of truth for a command's identity and metadata. The ribbon
    /// markup, tooltip callbacks, enable-state, licensing gate and generated docs all
    /// derive from this object — nothing is restated elsewhere, so they cannot drift.
    /// </summary>
    public sealed class CommandDefinition
    {
        /// <summary>Stable, unique key. Also the ribbon control id.</summary>
        public string Id { get; set; }

        /// <summary>Button caption shown on the ribbon.</summary>
        public string Label { get; set; }

        /// <summary>Bold tooltip title.</summary>
        public string Screentip { get; set; }

        /// <summary>One or two sentence tooltip description (Kutools/Ablebits quality).</summary>
        public string Supertip { get; set; }

        /// <summary>
        /// Base image name resolved by <c>IconProvider</c> (e.g. "ChangeCase" maps to the
        /// embedded "ChangeCaseIcon.png"). Null means no icon.
        /// </summary>
        public string ImageId { get; set; }

        /// <summary>
        /// Office built-in imageMso identifier used when no custom PNG is available.
        /// When set the ribbon emits imageMso="..." instead of getImage="GetImage".
        /// </summary>
        public string ImageMso { get; set; }

        /// <summary>Ribbon tab caption this command lives under.</summary>
        public string Tab { get; set; }

        /// <summary>Ribbon group caption within the tab.</summary>
        public string Group { get; set; }

        /// <summary>Sort order within the group (ascending).</summary>
        public int Order { get; set; }

        /// <summary>Default operating scope.</summary>
        public CommandScope Scope { get; set; }

        /// <summary>When true the command needs a non-empty range and is disabled otherwise.</summary>
        public bool RequiresSelection { get; set; }

        /// <summary>Undo capture strategy.</summary>
        public UndoMode UndoMode { get; set; }

        /// <summary>
        /// Licensing feature key checked by <c>CommandBase</c>. "core" is always available.
        /// Real tier gating is swapped in during the licensing phase without touching commands.
        /// </summary>
        public string LicenseFeature { get; set; }

        /// <summary>True to render as a large (32px) ribbon button, false for a small (16px) one.</summary>
        public bool LargeButton { get; set; }

        /// <summary>
        /// When set, this button is placed inside a dropdown &lt;menu&gt; with this label rather
        /// than appearing directly in the ribbon group.
        /// </summary>
        public string MenuParent { get; set; }

        /// <summary>Office imageMso shown on the parent menu dropdown button (only used when MenuParent is set).</summary>
        public string MenuParentImageMso { get; set; }

        /// <summary>
        /// Sort order of the parent menu container within its ribbon group.
        /// When 0, falls back to the minimum Order of the menu's child commands.
        /// </summary>
        public int MenuParentOrder { get; set; }

        public CommandDefinition()
        {
            Scope = CommandScope.Selection;
            RequiresSelection = true;
            UndoMode = UndoMode.None;
            LicenseFeature = "core";
            LargeButton = true;
            Order = 100;
        }
    }
}
