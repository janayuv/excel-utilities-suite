# Excel Utilities Suite — Architecture

## Overview

A **VSTO C# / .NET Framework 4.8** Excel add-in built on a dynamic Ribbon XML +
command-registry pattern. Every tool is a self-contained class; adding a new tool never
touches ribbon plumbing, tooltip callbacks, or error handling.

---

## Layer Map

```
Excel Process
└── VSTO Host (ThisAddIn.cs)
    ├── RibbonController        ← IRibbonExtensibility; XML built at runtime
    │   └── RibbonXmlBuilder    ← generates <customUI> from CommandRegistry
    ├── CommandRegistry         ← reflection-discovered [ExcelCommand] index
    │   └── Validate()          ← startup self-check; logs definition warnings
    ├── Services
    │   ├── UndoService         ← opt-in snapshot stack (20 deep, 200k-cell limit)
    │   ├── ProgressService     ← status-bar % updates
    │   ├── ErrorService        ← friendly dialog + rolling log (%APPDATA%)
    │   ├── SettingsService     ← per-user preferences (Properties.Settings)
    │   ├── LicenseService      ← state-machine stub (Trial/Licensed/Expired/…)
    │   ├── IconProvider        ← embedded PNG → IPictureDisp cache
    │   ├── KeyboardHook        ← WH_KEYBOARD_LL; Ctrl+Z → UndoService
    │   └── GridFocusService    ← row/column shape overlays on SelectionChange
    └── Commands
        ├── CommandBase         ← abstract; direct-action tools subclass this
        ├── DialogCommandBase   ← abstract; tools that open a WinForms dialog
        ├── OperationRunner     ← perf-toggle + undo capture + error guard
        └── Tools/              ← 22 files, 113 commands
```

---

## Key Design Decisions

### 1. Dynamic Ribbon XML (no Designer ribbon)

`RibbonXmlBuilder.Build()` is called once by `RibbonController.GetCustomUI()`.
It iterates every `CommandDefinition` in the registry, groups them by
`Tab → Group → Order`, and emits a `<customUI>` string.
The ribbon always matches the registry — they cannot drift.

**Tab consolidation** is a one-line dictionary in `TabRemap`: the 8 original feature
tabs collapse into **Utilities** and **Utilities +** (Kutools-style).

**imageMso fallback** in `ImageIdToMso`: every `ImageId` maps to a built-in Office icon
so all 113 tools have icons even before custom PNGs are provided.

### 2. Single Source of Truth — `CommandDefinition`

```csharp
public static readonly CommandDefinition Def = new CommandDefinition
{
    Id            = "Text.ChangeCase",
    Label         = "Change Case",
    Screentip     = "Change Case",
    Supertip      = "Convert selected cells to UPPER, lower, Title or Sentence case.",
    ImageId       = "ChangeCase",      // custom PNG  –OR–
    ImageMso      = "ChangeCase",      // built-in Office icon
    Tab           = "Editing",
    Group         = "Text",
    Order         = 10,
    Scope         = CommandScope.Selection,
    RequiresSelection = true,
    UndoMode      = UndoMode.FullSnapshot,
    LicenseFeature = "core"
};
```

The ribbon XML, every ribbon callback (`getLabel`, `getScreentip`, `getSupertip`,
`getImage`, `getEnabled`), the license gate, and the undo strategy all derive from
this object. Nothing is restated elsewhere.

### 3. Execution Pipeline (`OperationRunner.RunGuarded`)

```
CommandBase.Execute()
  → LicenseGate.Check(def)              blocks if feature locked
  → RangeResolver.Resolve(scope)        validates selection, returns target range
  → OperationRunner.RunGuarded(def, target, work)
      ├── StatusBarProgressReporter     start
      ├── UndoService.Begin(mode, target)   capture pre-state
      ├── app.ScreenUpdating = false
      ├── app.EnableEvents   = false
      ├── app.Calculation    = Manual
      ├── work(ctx)                     ← tool's Run() method
      ├── UndoService.Push(tx)          commit undo snapshot
      └── finally: restore all Excel state
```

On any exception: `tx.Restore()` (best-effort rollback), friendly error dialog,
log to `%APPDATA%\ExcelUtilities\error.log`.

### 4. Undo — Opt-In Per Command

| `UndoMode` | Behaviour |
|---|---|
| `None` | No capture (read-only, navigate, export tools) |
| `FullSnapshot` | Captures `Value2` of the whole target range up-front |
| `PartialSnapshot` | Cells recorded lazily via `ctx.Undo.RecordCell(cell)` during `Run` |
| `FormulaOnly` | Captures / restores `.Formula` only |

`FullSnapshot` is refused above **200,000 cells** (`FullSnapshotCellLimit`); the command
runs without undo and logs a warning. Stack depth cap: **20 transactions**.

**Ctrl+Z** is intercepted by `KeyboardHook` (`WH_KEYBOARD_LL`). When the add-in stack
has an entry it consumes the keypress and calls `UndoService.UndoLast()`; otherwise
the key passes through to Excel's native undo.

### 5. Icons — Two Paths

1. `def.ImageMso` set → `imageMso="..."` attribute in ribbon XML (no PNG needed).
2. `def.ImageId` + matching embedded PNG → `getImage="GetImage"` (custom asset).
3. Neither → button renders text-only.

Custom PNGs: `icon\{ImageId}Icon.png`, embedded with `LogicalName = "Icons.{name}"`.
`IconProvider` caches `Bitmap → IPictureDisp` conversions.

### 6. License Gate (Stub — Phase 5 swaps implementation)

State machine modelled now so the UI/command model needs no redesign later:

```
Trial ──────────────────► Licensed
  └─► Expired ──────────► (re-activate) ─► Licensed
            └─► Offline grace ─► Expired
FeatureLocked  (tool tier not unlocked)
```

`CommandBase` calls `LicenseGate.Check(def)`. Locked tools show in the ribbon
(enabled), prompt to upgrade on click. The stub returns `Licensed` for all features.

### 7. Grid Focus

`GridFocusService.OnSelectionChange()` runs on every `SheetSelectionChange`.
Deletes `__GF_Row__` / `__GF_Col__` shapes, then places new semi-transparent
rectangles over the active cell's row and column. Configurable via `GridFocusDialog`
(shape, style, colour, transparency). Shapes are sent to back and marked non-locked.

### 8. Startup Self-Check

`CommandRegistry.Validate()` runs after `Initialize()`. It logs a `WARN` for every
command missing `Label`, `Tab`, `Group`, or tooltip text. A clean startup logs:

```
[Self-check] PASS — 113 commands registered, 0 definition warnings.
```

---

## File Layout

```
D:\Excel Addins\
├── Utilities\
│   ├── ThisAddIn.cs
│   ├── Commands\
│   │   ├── CommandBase.cs
│   │   ├── CommandContext.cs
│   │   ├── CommandModel.cs          CommandDefinition, UndoMode, scope enums
│   │   ├── CommandRegistry.cs       reflection discovery + self-check
│   │   ├── IExcelCommand.cs
│   │   ├── OperationRunner.cs
│   │   └── Tools\                   22 files, 113 [ExcelCommand] classes
│   ├── Dialogs\
│   │   └── DialogBase.cs
│   ├── Ribbon\
│   │   ├── RibbonController.cs
│   │   └── RibbonXmlBuilder.cs
│   └── Services\
│       ├── ErrorService.cs
│       ├── GridFocusService.cs
│       ├── IconProvider.cs
│       ├── KeyboardHook.cs
│       ├── LicenseService.cs
│       ├── ProgressService.cs
│       ├── SettingsService.cs
│       └── UndoService.cs
├── icon\                            PNG icon assets
├── deploy\                          Inno Setup + VstoAddinInstaller
└── docs\
    ├── architecture.md              this file
    ├── catalog.md                   full command catalog (113 tools)
    └── testing-matrix.md            test plan and pass/fail records
```

---

## Adding a New Tool (checklist)

1. Create `Commands/Tools/MyCommand.cs` with `[ExcelCommand]` class.
2. Declare `CommandDefinition` — Id, Label, Screentip, Supertip, ImageMso, Tab, Group, Order, UndoMode.
3. Implement `Run(CommandContext ctx)` (direct) or `CreateDialog()` (dialog).
4. Add `<Compile Include="Commands\Tools\MyCommand.cs" />` to `utilities.csproj`.
5. Rebuild — ribbon button, tooltip, icon, undo appear automatically.
6. Run the Testing Matrix against the new tool.
