# Excel Utilities Suite — CLAUDE.md

> Production reference for Claude Code sessions. Read this before touching any file.

## Project at a Glance

- **Type**: VSTO C# Excel add-in, .NET Framework 4.8, targets Excel 2016+
- **Main project**: `Utilities\utilities.csproj` (the root-level `utilities.csproj` is a legacy stub)
- **Solution**: `utilities.sln`
- **~113 tools** across 24 command files, discovered by reflection at startup
- **No hot-reload** — rebuild required for every code change
- **No test project** — manual testing via `docs/testing-matrix.md`

---

## Folder Structure

```
D:\Excel Addins\
├── utilities.sln
├── Utilities\                      main add-in project
│   ├── utilities.csproj            authoritative project file
│   ├── ThisAddIn.cs                VSTO entry point
│   ├── ExcelHelper.cs              *** LEGACY — pre-framework utility; do not copy patterns ***
│   ├── DataConverterHelper.cs      *** LEGACY — pre-framework utility; do not copy patterns ***
│   ├── Commands\
│   │   ├── CommandModel.cs         CommandDefinition, UndoMode, CommandScope, [ExcelCommand]
│   │   ├── IExcelCommand.cs        interface: Definition + Execute()
│   │   ├── CommandBase.cs          abstract bases: CommandBase / DialogCommandBase
│   │   ├── CommandContext.cs       per-invocation context passed to Run()
│   │   ├── CommandRegistry.cs      reflection discovery + startup self-check
│   │   ├── OperationRunner.cs      single execution guard for all tools
│   │   └── Tools\                  24 files, ~113 [ExcelCommand] classes
│   ├── Dialogs\
│   │   ├── DialogBase.cs           WinForms base; routes mutations through OperationRunner
│   │   └── FindRunDialog.cs        searchable command picker (Find & Run)
│   ├── Ribbon\
│   │   ├── RibbonController.cs     IRibbonExtensibility; callback dispatch hub
│   │   └── RibbonXmlBuilder.cs     generates <customUI> XML from registry at runtime
│   └── Services\
│       ├── ErrorService.cs         friendly dialog + rolling log (%APPDATA%)
│       ├── GridFocusService.cs     row/col shape overlays on SheetSelectionChange
│       ├── IconProvider.cs         embedded PNG → IPictureDisp cache, dark-mode detection
│       ├── KeyboardHook.cs         WH_KEYBOARD_LL Ctrl+Z interception
│       ├── LicenseService.cs       ILicenseService, StubLicenseService, License.Current
│       ├── LicenseSalt.cs          *** SECRET — gitignored, never read or edit ***
│       ├── LicenseSalt.example.cs  template for the above
│       ├── ProgressService.cs      IProgressReporter, StatusBarProgressReporter
│       ├── RealLicenseService.cs   offline HMAC-SHA256 key validation
│       ├── RepeatService.cs        last-action record/replay ("Repeat Last Tool")
│       ├── SettingsService.cs      per-user key=value prefs (%APPDATA%)
│       └── UndoService.cs          UndoTransaction, IUndoRecorder, 20-deep stack
├── KeyGen\                         internal key-generator tool (not shipped to users)
├── icon\                           32×32 light PNGs   → LogicalName: Icons.{Name}Icon.png
├── icon-16\                        16×16 light PNGs   → LogicalName: Icons16.{Name}Icon.png
├── icon-dark\                      32×32 dark PNGs    → LogicalName: IconsDark.{Name}Icon.png
├── icon-dark-16\                   16×16 dark PNGs    → LogicalName: IconsDark16.{Name}Icon.png
├── deploy\                         VstoAddinInstaller + Inno Setup scripts
└── docs\
    ├── architecture.md             layer map + design decisions
    ├── catalog.md                  full command catalog (113 tools)
    └── testing-matrix.md           manual test plan + pass/fail records
```

---

## Architecture Overview

```
Excel Process
└── VSTO Host (ThisAddIn.cs)
    ├── RibbonController        ← IRibbonExtensibility; XML built at runtime from registry
    │   └── RibbonXmlBuilder    ← emits <customUI> grouping commands by Tab/Group/Order
    ├── CommandRegistry         ← reflection-discovers [ExcelCommand] classes at startup
    │   └── Validate()          ← self-check; logs WARN for missing Label/Tab/Group/tooltip
    ├── Services\               ← stateless/static helpers (no DI framework)
    │   ├── UndoService         ← opt-in snapshot stack, 20 deep, 200k-cell cap
    │   ├── RepeatService       ← captures last RunGuarded closure; replays on demand
    │   ├── ErrorService        ← friendly dialog + rolling log (rolls at 1 MB)
    │   ├── SettingsService     ← key=value file under %APPDATA%\ExcelUtilitiesSuite
    │   ├── LicenseService      ← state machine: Trial/Licensed/Expired/Offline/FeatureLocked
    │   ├── IconProvider        ← embedded PNG → IPictureDisp; dark-mode via registry
    │   ├── KeyboardHook        ← WH_KEYBOARD_LL; Ctrl+Z → UndoService (posted to UI thread)
    │   └── GridFocusService    ← shape overlays on SheetSelectionChange
    └── Commands\
        ├── CommandBase         ← abstract; direct-action tools subclass this
        ├── DialogCommandBase   ← abstract; dialog tools subclass this
        ├── OperationRunner     ← perf-toggle + undo + error guard + RepeatService capture
        └── Tools\              ← concrete [ExcelCommand] classes; ~4–6 per file by feature area
```

---

## Command System

### The Interface

```csharp
public interface IExcelCommand
{
    CommandDefinition Definition { get; }
    void Execute();   // called by the ribbon; implemented by CommandBase / DialogCommandBase
}
```

### CommandDefinition — Single Source of Truth

Every command declares one `public static readonly CommandDefinition Def`. Nothing is stated twice:
the ribbon XML, every ribbon callback, the license gate, and undo strategy all derive from it.

| Field | Purpose |
|-------|---------|
| `Id` | Stable unique key and ribbon control tag. Format: `"Category.PascalName"` (e.g. `"Text.ChangeCase"`) |
| `Label` | Ribbon button caption |
| `Screentip` | Bold tooltip title |
| `Supertip` | 1–2 sentence tooltip body (Kutools/Ablebits quality standard) |
| `ImageId` | Custom PNG name (e.g. `"ChangeCase"` resolves to `ChangeCaseIcon.png` embedded resource) |
| `ImageMso` | Office built-in icon name — used instead of `ImageId` when no custom PNG exists |
| `Tab` | Logical tab name; `TabRemap` in `RibbonXmlBuilder` collapses to "Utilities" / "Utilities +" |
| `Group` | Ribbon group within the tab |
| `Order` | Sort position within the group (ascending); default 100 |
| `Scope` | `Selection` / `Worksheet` / `Workbook` — controls what `RangeResolver` returns |
| `RequiresSelection` | When true, `Execute()` aborts with a warning if no range is resolved |
| `UndoMode` | See Undo System section |
| `LicenseFeature` | Feature key checked at gate; `"core"` is always available |
| `LargeButton` | `true` = 32 px ribbon button; `false` = 16 px |
| `MenuParent` | When set, button lives inside a `<menu>` dropdown with this label |

> **`CommandScope.Workbook` does not iterate all sheets.** Both `CommandScope.Worksheet` and
> `CommandScope.Workbook` resolve to `ActiveSheet.UsedRange` via the same code path in
> `RangeResolver`. A command declaring `Scope = Workbook` operates on the active sheet only.
> Any genuine workbook-wide command must manually enumerate `app.Worksheets` inside `Run`.

### Direct-Action Commands

Subclass `CommandBase`, implement `Run(CommandContext ctx)`:

```csharp
[ExcelCommand]
public sealed class TrimTextCommand : CommandBase
{
    public static readonly CommandDefinition Def = new CommandDefinition
    {
        Id        = "Text.TrimText",
        Label     = "Trim Text",
        Screentip = "Trim Text",
        Supertip  = "Remove leading, trailing and excess internal spaces from selected cells.",
        ImageId   = "TrimText",
        Tab       = "Editing",
        Group     = "Text",
        Order     = 10,
        Scope     = CommandScope.Selection,
        UndoMode  = UndoMode.FullSnapshot,
    };

    public override CommandDefinition Definition { get { return Def; } }

    protected override void Run(CommandContext ctx)
    {
        foreach (Excel.Range area in ctx.Target.Areas)
        {
            object[,] values = area.Value2 as object[,];
            if (values == null) // single-cell selection — Value2 is a scalar, not an array
            {
                if (area.Value2 is string s) area.Value2 = s.Trim();
                continue;
            }
            int rows = values.GetLength(0), cols = values.GetLength(1);
            for (int r = 1; r <= rows; r++)        // Value2 arrays are 1-based, not 0-based
                for (int c = 1; c <= cols; c++)
                    if (values[r, c] is string sv) values[r, c] = sv.Trim();
            area.Value2 = values;                   // one COM write for the whole area
        }
    }
}
```

### Dialog Commands

Subclass `DialogCommandBase`, implement `CreateDialog()`. The dialog subclasses `DialogBase` and calls
`RunOperation(def, target, work)` to route the mutation through `OperationRunner`.

### DialogBase API

Every dialog routes mutations through `OperationRunner` by subclassing `DialogBase`. Key members:

| Member | Purpose |
|--------|---------|
| `WireButtons(Button accept, Button cancel)` | Wires Enter → OK, Esc → Cancel; call in constructor |
| `SetError(Control control, string message)` | Displays inline validation error next to a control |
| `HasNoErrors(params Control[] controls)` | Returns true only if all specified controls are error-free |
| `CurrentSelection` | Returns the live `Excel.Range` — do not cache; re-read on each use |
| `RunOperation(def, target, work)` | Routes the mutation lambda through `OperationRunner.RunGuarded` |

**Defaults applied by `DialogBase`:** `FormBorderStyle = FixedDialog`, `AutoScaleMode = Font`,
`KeyPreview = true`, `MaximizeBox = MinimizeBox = false`. Override `FormBorderStyle` to `Sizable`
only when the dialog explicitly needs resize (as `FindRunDialog` does).

**Validation pattern:** call `SetError` to flag invalid controls, then guard `RunOperation` behind `HasNoErrors`:

```csharp
private void OnOkClick(object sender, EventArgs e)
{
    SetError(_modeCombo, null);
    if (_modeCombo.SelectedIndex < 0) SetError(_modeCombo, "Select a mode.");
    if (!HasNoErrors(_modeCombo)) return;
    RunOperation(MyCommand.Def, CurrentSelection, ctx => { /* mutation */ });
}
```

```csharp
[ExcelCommand]
public sealed class ChangeCaseCommand : DialogCommandBase
{
    public static readonly CommandDefinition Def = new CommandDefinition { /* ... */ };
    public override CommandDefinition Definition { get { return Def; } }
    protected override DialogBase CreateDialog() { return new ChangeCaseDialog(); }
}
```

---

## Registry Pattern

`CommandRegistry.Initialize()` (called once at startup — safe to call multiple times):

1. Reflects over `Assembly.GetExecutingAssembly()`
2. Finds every non-abstract type with `[ExcelCommand]` attribute that implements `IExcelCommand`
3. Requires a parameterless constructor — logs WARN and skips if missing
4. Instantiates, reads `Definition.Id`, checks for duplicates
5. Calls `Validate()` to log WARN for any definition missing Label, Tab, Group, or tooltip

**A clean startup logs:**
```
[Self-check] PASS — 113 commands registered, 0 definition warnings.
```

`GetCustomUI()` in `RibbonController` calls `CommandRegistry.Initialize()` defensively because Excel
calls `GetCustomUI` before `ThisAddIn_Startup` fires.

---

## Ribbon Architecture

`RibbonXmlBuilder.Build()` generates the entire `<customUI>` string from the registry at load time.
No Designer ribbon. The ribbon always matches command definitions — they cannot drift.

**Tab consolidation** (`TabRemap` dictionary):

| Source Tab | Rendered Tab |
|---|---|
| Editing, Insert, Select & Navigate | **Utilities** |
| Formula & Statistics, Data & Cleaning, Export / Import, Workbook & Sheets, Printing | **Utilities +** |

A fixed **Suite** tab is appended last, containing:
- **Quick** group: Find & Run (`sys.findrun`), Repeat Last Tool (`sys.repeat`)
- **License** group: Activate / Deactivate (hidden when licensed)
- **History** group: Undo Last Action (`sys.undo`)
- **Help** group: About, Open Log

**Ribbon callbacks** are dispatched through `RibbonController` by `control.Tag` (= command Id).
System controls use hardcoded tag constants (`sys.undo`, `sys.repeat`, `sys.findrun`, `sys.about`, `sys.openlog`).

**Ribbon refresh**: Buttons with dynamic labels or enabled state use `getLabel` and `getEnabled`
callbacks. The ribbon only re-queries these when `_ribbon.Invalidate()` is called.
`RibbonController.Invalidate()` is called automatically after every `OnAction` dispatch. If you add
a new dynamic button whose state changes on selection change, ensure `Invalidate()` is also called
from `ThisAddIn.OnSheetSelectionChange`.

**Icon resolution** (two paths):
1. `def.ImageId` set → `getImage="GetImage"` → `IconProvider.GetPicture(imageId)` → embedded PNG
2. `def.ImageMso` set → `imageMso="..."` attribute in XML → Office built-in icon
3. Both set → `ImageMso` wins (emitted as attribute; no `getImage` callback)

---

## Undo System

### UndoMode (per-command)

| Mode | Behaviour | When to use |
|------|-----------|-------------|
| `None` | No capture; command runs without undo | Read-only, navigate, select, export tools |
| `FullSnapshot` | Captures `Value2` of entire target range up-front | Small ranges; refused above **200,000 cells** |
| `PartialSnapshot` | Cells recorded lazily via `ctx.Undo.RecordCell(cell)` inside `Run` | Sparse mutations over large ranges |
| `FormulaOnly` | Captures/restores `.Formula` only | Formula-rewriting tools |

### Stack

- `UndoService` maintains a `Stack<UndoTransaction>` capped at **20** entries
- `UndoService.Begin()` creates and up-front-captures the transaction; returned as `IUndoRecorder`
- `UndoService.Push(tx)` commits on success
- `tx.Restore()` is called on failure (best-effort, wrapped in try/catch)
- `UndoService.UndoLast()` pops and restores; ribbon Undo button is dynamically labelled

### Ctrl+Z Interception

`KeyboardHook` (WH_KEYBOARD_LL) intercepts Ctrl+Z at the OS level:
- If `UndoService.CanUndo` → consume the keystroke, post undo work to `SynchronizationContext`
- Otherwise → pass through to Excel's native undo via `CallNextHookEx`

**Critical rule**: The `HookCallback` must complete in < 200 ms and must never do I/O, COM calls,
or anything blocking. Only the pure in-memory check `UndoService.CanUndo` is safe inside it.
All real work is posted to the captured UI-thread `SynchronizationContext`.

---

## OperationRunner

The single choke-point for every tool mutation. Both `CommandBase.Execute()` and
`DialogBase.RunOperation()` route through here so behaviour is identical everywhere.

```
RunGuarded(def, target, work):
  1. Create StatusBarProgressReporter
  2. UndoService.Begin(mode, target)      — capture pre-state
  3. Snapshot app.ScreenUpdating / EnableEvents / Calculation
  4. app.ScreenUpdating = false
     app.EnableEvents   = false
     app.Calculation    = Manual
  5. work(ctx)                             — tool's Run() or dialog lambda
  6. UndoService.Push(tx)                  — commit undo snapshot
  7. RepeatService.Record(def, work)       — capture for "Repeat Last Tool"
  finally:
     Restore ScreenUpdating / EnableEvents / Calculation (always, even on exception)
     progress.Done()
  catch:
     tx.Restore()                          — best-effort rollback
     ErrorService.Handle(ex, def)          — friendly dialog + log
```

**Never** call `OperationRunner.RunGuarded` from inside a `RunGuarded` work closure.

---

## RepeatService

Records the last successful tool invocation and replays it against the current selection.

**Record** happens in `OperationRunner.RunGuarded` on success. Non-repeatable tools are silently
skipped (previous record stays intact):

| Excluded groups | Excluded IDs |
|----------------|-------------|
| Printing, Export, Import | `Range.SwapRanges`, `Range.CompareRanges`, `Sheet.MergeWorkbooks`, `Sheet.SplitByColumn` |

**Replay** (`RepeatService.Replay()`):
1. Re-resolves the current selection via `RangeResolver.Resolve`
2. Calls `OperationRunner.RunGuarded(def, target, work)` with the captured closure

The captured `work` closure already has the user's dialog settings baked in. Replaying it re-applies
the same operation without re-opening any dialog. The replay is recorded as the new last action —
safe because the closure calls `RunGuarded`, not `Replay`.

**Find & Run** (`FindRunDialog`): searchable picker over all registered commands. Filters by label,
id, or tooltip. Pins the last-used tool to the top when no search is active. Chosen command is
dispatched via `RibbonController.ShowFindRun`, which calls `cmd.Execute()` — flows through
`RunGuarded` and updates the Repeat pointer automatically.

---

## Excel Interop Rules

These rules apply to every tool's `Run()` method and dialog lambdas:

1. **Always work through `ctx.Target`** — the range already resolved from the declared `CommandScope`.
   Never call `Globals.ThisAddIn.Application.Selection` inside `Run`.
   When `RequiresSelection = false`, **`ctx.Target` will be null** — `RangeResolver` returns null
   and `Execute()` skips the selection guard. Always null-check `ctx.Target` in commands that
   declare `RequiresSelection = false`.

2. **`ScreenUpdating`, `EnableEvents`, `Calculation`** — already set to performance-safe values by
   `OperationRunner` before `Run` is called. Do not toggle them again inside `Run`.

3. **Use `Value2` not `Value`** — `Value` triggers format parsing and is slower. `Value2` returns
   the raw double/string/bool. Cells containing Excel errors (`#N/A`, `#REF!`, `#VALUE!`, etc.)
   return a COM error integer code via `Value2`, not a string — `value is string` checks silently
   skip error cells. This is usually the correct behaviour for text-processing tools.

4. **Use `.Formula` for formula-rewriting tools**, paired with `UndoMode.FormulaOnly`. Always use
   `.Formula` (English function names), never `.FormulaLocal` — `.FormulaLocal` uses
   locale-specific names (e.g. `SUMME` on German Excel) and breaks portability.

5. **Bulk read/write is the standard pattern for mutating tools.** Read the entire area in one COM
   call, modify the managed array, then write it back in one COM call. Three critical invariants:

   - **`Value2` arrays are 1-based**: `values[1, 1]` is the first cell, not `values[0, 0]`.
     Use `values.GetLength(0)` / `values.GetLength(1)` for row/column counts.
   - **Single-cell ranges return a scalar**: `area.Value2 as object[,]` returns `null` when the
     target is a single cell — `Value2` is a plain `object` in that case. Always null-check.
   - **Always iterate `ctx.Target.Areas`**: `CommandScope.Selection` can include non-contiguous
     multi-area selections (Ctrl+click). Apply the bulk pattern to each area independently.

   ```csharp
   foreach (Excel.Range area in ctx.Target.Areas)
   {
       object[,] values = area.Value2 as object[,];
       if (values == null) // single cell — Value2 is scalar, not an array
       {
           // handle single-cell case inline
           continue;
       }
       int rows = values.GetLength(0), cols = values.GetLength(1);
       for (int r = 1; r <= rows; r++)   // 1-based — not 0-based
           for (int c = 1; c <= cols; c++)
               // mutate values[r, c]
       area.Value2 = values; // one COM write for the whole area
   }
   ```

   Use `foreach (Excel.Range cell in area.Cells)` only for `PartialSnapshot` tools calling
   `ctx.Undo.RecordCell(cell)` per modified cell, or for genuinely sparse conditional mutations.
   Never use it for bulk reads or writes.

6. **Progress reporting** — call `ctx.Progress.Report(fraction)` for any loop over a large range.
   Throttled to whole-percent changes internally; no-op for small ranges is fine.

7. **`PartialSnapshot` pattern** — call `ctx.Undo.RecordCell(cell)` _before_ writing to each cell
   the first time. The recorder only saves the first value per address.

8. **`MessageBox.Show` is acceptable inside `Run` for pre-condition failures** (e.g., "Select at
   least two rows.") — it is a pure WinForms call unaffected by `EnableEvents = false`. **Never**
   open dialogs that interact with Excel (`Application.GetOpenFilename`,
   `Application.GetSaveAsFilename`) inside `Run` — those require Excel to be interactive. For
   unexpected runtime errors, throw an exception; `OperationRunner` catches it and shows the error
   dialog.

9. **Status bar is best-effort** — `StatusBarProgressReporter` swallows all exceptions. Never
   depend on it for control flow.

---

## Performance Rules

Excel add-ins degrade quickly when they issue many small COM calls. Follow these rules for every
tool that operates on ranges larger than a handful of cells:

- **Read entire ranges in one COM call**: `object[,] values = range.Value2;` — never read cells one-by-one inside a loop.
- **Write entire ranges in one COM call**: assign the modified array back — `range.Value2 = values;` — never write cells one-by-one.
- **Never call `WorksheetFunction` inside a loop** — each call crosses the COM boundary. Pre-compute results in managed code instead.
- **Prefer managed arrays over COM access** — work on `object[,]` in C#; touch COM only at the read boundary and the write boundary.
- **Avoid repeated property reads on the same COM object** — cache `cell.Row`, `range.Rows.Count`, etc. into local variables before entering a loop.
- **`OperationRunner` already disables `ScreenUpdating`, `EnableEvents`, and `Calculation`** — do not undo this inside `Run`.
- **Test on 50,000+ row datasets** before marking a tool complete. Status-bar % updates (via `ctx.Progress.Report`) are the only acceptable overhead inside a large loop.
- **`PartialSnapshot` undo is required for tools that mutate sparse cells** across a large range — `FullSnapshot` is refused above 200,000 cells.

---

## COM Object Handling Rules

VSTO uses embedded interop types (`EmbedInteropTypes = true`), so `Marshal.ReleaseComObject` is
generally not required. Follow these rules instead:

1. **Do not hold Excel COM references across event handlers or async boundaries.**
   Pattern: `GridFocusService._lastSheet` is cleared in `OnWorkbookBeforeClose` via `ClearLastSheet()`.
   Apply the same pattern to any service that caches a worksheet or range.

2. **COM calls must happen on the Excel UI thread (STA).** Never call Excel interop from a background
   thread. The `KeyboardHook` callback demonstrates the correct pattern: only pure in-memory work
   inside the callback; real COM work posted to `SynchronizationContext`.

3. **Wrap COM enumeration in try/catch.** `ws.Shapes` iteration and similar COM collections can throw
   if the sheet is being modified concurrently. Swallow at the loop level.

4. **`app.ActiveSheet as Excel.Worksheet` may return null.** Some hook callbacks fire on chart sheets
   or unsupported sheet types. Always null-check.

5. **`app.ActiveWorkbook` may be null.** Guard all workbook-level access with a null check.
   `RangeResolver.Resolve` already does this — use it instead of querying `ActiveWorkbook` directly.

6. **Do not cache `Excel.Range` objects** as long-lived fields. Ranges are COM wrappers; the
   underlying cell moves when rows/columns are inserted. Re-resolve from address when needed.

---

## Naming Conventions

| Thing | Convention | Example |
|---|---|---|
| Command class | `{Feature}Command` (sealed) | `TrimTextCommand` |
| Dialog class | `{Feature}Dialog` (internal sealed) | `TrimTextDialog` |
| `CommandDefinition.Id` | `"Category.PascalName"` | `"Text.TrimText"` |
| `CommandDefinition.Tab` | Title case; maps via `TabRemap` | `"Editing"` |
| `CommandDefinition.Group` | Title case | `"Text"` |
| `ImageId` | PascalName matching the PNG file base | `"TrimText"` → `TrimTextIcon.png` |
| `SettingsService` keys | `"CommandId.SettingName"` | `"ChangeCase.LastMode"` |
| Tool file | Feature area + `Commands.cs` | `TextCommands.cs`, `RangeCommands.cs` |

Command ID uniqueness is enforced at startup. Duplicates are logged and the second is skipped.

---

## Build Commands

All commands run from the repo root (`D:\Excel Addins`).

```powershell
# Debug build (default for development)
msbuild "Utilities\utilities.csproj" /p:Configuration=Debug /nologo /v:minimal

# Release build
msbuild "Utilities\utilities.csproj" /p:Configuration=Release /nologo /v:minimal

# Full solution (includes KeyGen and deploy demo projects)
msbuild utilities.sln /p:Configuration=Release /nologo /v:minimal

# Errors and warnings only
msbuild "Utilities\utilities.csproj" /p:Configuration=Debug /nologo /v:quiet
```

Output paths:
- Debug: `Utilities\bin\Debug\utilities.dll`
- Release: `Utilities\bin\Release\utilities.dll`

**There is no hot-reload.** Close Excel, rebuild, reopen Excel. The VSTO runtime reloads the add-in
on Excel startup.

Definition warnings appear in `%APPDATA%\ExcelUtilitiesSuite\suite.log` (Suite → Help → Open Log).

---

## Adding a New Tool — Checklist

1. **Create or append to** `Utilities\Commands\Tools\<Category>Commands.cs`
   - Annotate with `[ExcelCommand]`; class must be `sealed` with a parameterless constructor
   - Declare `public static readonly CommandDefinition Def = new CommandDefinition { ... }`
   - Direct tool: subclass `CommandBase`, implement `Run(CommandContext ctx)`
   - Dialog tool: subclass `DialogCommandBase`, implement `CreateDialog()`; dialog class is `internal sealed` in the same file

2. **Register in `Utilities\utilities.csproj`** (only if the file is new):
   ```xml
   <Compile Include="Commands\Tools\<Category>Commands.cs" />
   ```

3. **Add icons** to all four icon folders (`icon\`, `icon-16\`, `icon-dark\`, `icon-dark-16\`)
   and embed each in csproj:
   ```xml
   <EmbeddedResource Include="..\icon\{Name}Icon.png">
     <Link>Icons\{Name}Icon.png</Link>
     <LogicalName>Icons.{Name}Icon.png</LogicalName>
   </EmbeddedResource>
   ```
   If no custom PNG, use `ImageMso` in `CommandDefinition` instead.

4. **Rebuild** — ribbon button, tooltip, icon, undo, license gate appear automatically.

5. **Run the Testing Matrix** against the new tool.

---

## Adding New Commands

Use this workflow for every new command, regardless of size. The discipline pays off once the suite
exceeds 50–100 tools and overlap becomes invisible.

### Before implementing

1. **Search existing commands for overlap** — grep for the feature name across `Commands/Tools/`.
   Duplicate tools create user confusion and inflate the ribbon. If partial overlap exists, extend
   the existing command with a new mode rather than adding a second button.
2. **Verify category placement** — confirm the correct `Tab` and `Group` using the `TabRemap`
   dictionary in `RibbonXmlBuilder` and the existing group names in `GroupImageMso`. Place the
   command alongside semantically related tools, not just the most convenient file.
3. **Define `UndoMode`** — choose before writing a single line of `Run`:
   - `None` for read-only, navigate, select, export
   - `FullSnapshot` for small destructive edits (< 200k cells)
   - `PartialSnapshot` for sparse mutations over large ranges
   - `FormulaOnly` for formula-rewriting
4. **Define repeatability** — will this command be safe to re-fire silently via Repeat Last Tool?
   If not (file dialog, multi-range, printing), add its `Id` to `NonRepeatableIds` or ensure
   its `Group` matches `NonRepeatableGroups` in `RepeatService`.
5. **Define ribbon location** — `Tab`, `Group`, `Order`, `LargeButton`, and optionally `MenuParent`.
   Check `Order` values of neighbours in the same group to avoid collisions.
6. **Define icon** — either a custom `ImageId` with four PNG sizes (32/16 light/dark) or an
   `ImageMso` Office built-in. Do not ship a command with no icon.
7. **Define tooltip** — both `Screentip` (bold title) and `Supertip` (1–2 sentence description at
   Kutools/Ablebits quality). The self-check at startup will WARN if either is missing.
8. **Define command ID** — format `"Category.PascalName"` (e.g. `"Data.FillDownBlanks"`). IDs are
   permanent; changing one after release silently breaks the ribbon for existing users.

### After implementing

1. **Verify registry discovery** — open the log (`Suite → Help → Open Log`) after reloading Excel
   and confirm `[Self-check] PASS — N commands registered, 0 definition warnings`.
2. **Verify ribbon appearance** — confirm the button appears in the correct tab, group, and position
   with the correct label and icon in both light and dark Office themes.
3. **Verify Repeat Last Tool** — run the command, then click Suite → Repeat. Confirm it re-applies
   to a new selection. If it should be non-repeatable, confirm it does not update the Repeat pointer.
4. **Verify Undo** — run the command, then press Ctrl+Z or click Suite → Undo. Confirm cells revert
   to their original values. For `UndoMode.None` tools, confirm Ctrl+Z passes through to Excel.
5. **Verify large-range performance** — run against a 50,000+ row selection. Confirm it completes in
   a reasonable time and that the status bar shows progress. No cell-by-cell COM calls.

---

## Command Design Checklist

Every command must have all of the following defined before it is considered complete:

| Requirement | Field / Location |
|---|---|
| Category | `CommandDefinition.Tab` + `CommandDefinition.Group` |
| Command ID | `CommandDefinition.Id` — format `"Category.PascalName"`, permanent |
| Ribbon label | `CommandDefinition.Label` |
| Screentip (bold tooltip title) | `CommandDefinition.Screentip` |
| Supertip (tooltip body) | `CommandDefinition.Supertip` — 1–2 sentences |
| Undo mode | `CommandDefinition.UndoMode` — never leave as default `None` for mutating tools |
| Repeat support | Confirm via `RepeatService.IsRepeatable`; add to exclusion list if unsafe |
| Progress reporting | `ctx.Progress.Report(fraction)` called inside any loop over > ~1,000 cells |
| Error handling | Mutations inside `Run` only — `OperationRunner` handles all exceptions |
| Cancellation behaviour | Document whether the tool honours `ctx.Progress.IsCancellationRequested` |

The `CommandRegistry.Validate()` self-check enforces Label, Tab, Group, and tooltip at startup.
The rest of this checklist is enforced by code review.

---

## Release Process

1. Ensure `LicenseSalt.cs` is present (not the example file) and correct.
2. Build Release configuration.
3. In Visual Studio: right-click `Utilities` → Publish → follow ClickOnce wizard.
   Published output goes to `Utilities\12\` by default.
4. Run Inno Setup Compiler on `deploy\VstoAddinInstaller\` to produce the standalone installer `.exe`.
5. Increment `ApplicationVersion` in `Utilities\utilities.csproj` before each publish.

---

## Security-Sensitive Files

| File | Rule |
|------|------|
| `Utilities\Services\LicenseSalt.cs` | **Never read, edit, or commit.** Gitignored. Contains the HMAC secret. |
| `Utilities\Services\LicenseSalt.example.cs` | Template only — shows the required class shape with a placeholder value. |
| `KeyGen\` | Internal keygen. Not shipped. Uses `RealLicenseService.GenerateKey()`. |

### License Key System

- Format: `XXXXX-XXXXX-XXXXX-XXXXX` (20 hex chars, dash-separated in groups of 5)
- Validation: `HMAC-SHA256((machineId + body).ToUpperInvariant(), salt)` — last 8 hex chars = first 4 bytes of hash
- Machine ID: SHA256 of `(drive root + machine name)`, first 8 bytes as hex
- States: `Trial(30d)` → `Licensed` | `Expired`; machine change → `Offline(14d grace)` → `Expired`
- Registry root: `HKCU\Software\ExcelUtilitiesSuite`

---

## Runtime Paths

| Path | Purpose |
|------|---------|
| `%APPDATA%\ExcelUtilitiesSuite\suite.log` | Rolling log (1 MB cap, one `.1` backup) |
| `%APPDATA%\ExcelUtilitiesSuite\settings.txt` | Per-user tool preferences (`CommandId.Key=value`) |
| `HKCU\Software\ExcelUtilitiesSuite` | Trial start date, license key, machine fingerprint |

---

## Claude Code Hard Rules

These rules prevent architectural drift. They apply unconditionally in every session.

**Do not:**

- Create alternative command architectures (e.g. a second `ICommand` interface, a separate plugin
  system, a factory pattern layered on top of the registry). All tools use `[ExcelCommand]` +
  `CommandBase` or `DialogCommandBase`. Period.
- Bypass `CommandRegistry` — do not instantiate or invoke commands directly outside of the registry
  dispatch path. The ribbon, Find & Run, and Repeat Last Tool all depend on it.
- Add ribbon logic into command classes — commands know nothing about the ribbon. Tab, group, order,
  label, and icon all live in `CommandDefinition` and are consumed exclusively by `RibbonXmlBuilder`.
- Add business logic into ribbon classes — `RibbonController` and `RibbonXmlBuilder` dispatch and
  generate XML only. No cell manipulation, no range logic, no service calls beyond `CommandRegistry.Get`.
- Introduce service locators or IoC/DI containers — all services are static classes. This is
  intentional for a .NET Framework 4.8 add-in with no package manager. Do not introduce Unity,
  Autofac, Microsoft.Extensions.DependencyInjection, or any equivalent.
- Modify licensing code (`LicenseService.cs`, `RealLicenseService.cs`, `LicenseSalt.cs`,
  `RealLicenseService.ValidateKey`, `RealLicenseService.GenerateKey`) without an explicit instruction
  from the user. License logic is security-critical and offline — mistakes cannot be patched without
  a new release.
- Modify deployment configuration (`deploy\`, `utilities.csproj` `<PublishUrl>`,
  `ApplicationVersion`, ClickOnce manifests) without an explicit instruction from the user.
- Manually edit any `*.Designer.cs` file — `ThisAddIn.Designer.cs`, form designer files, and any
  other `*.Designer.cs` files are Visual Studio–generated and will be silently overwritten on the
  next designer pass.

---

## Recommended Agents

These agents are worth creating for this repository. They provide project-specific expertise that
generic agents cannot match. Create them in `.claude/agents/`.

| Agent | Value | Focus |
|---|---|---|
| `excel-interop-reviewer` | **Very High** | Flags cell-by-cell COM loops, missing `Value2` bulk reads/writes, forgotten `ScreenUpdating` restore, COM references held across event handlers |
| `command-planner` | **Very High** | Reviews a new-command brief against the Adding New Commands checklist — checks for overlap, confirms UndoMode, repeatability, ribbon placement, tooltip quality, and ID format before any code is written |
| `ribbon-reviewer` | **High** | Validates `CommandDefinition` completeness (`Tab`, `Group`, `Order`, `ImageId`/`ImageMso`, `Screentip`, `Supertip`), confirms `TabRemap` placement, checks for `Order` collisions within a group |
| `release-reviewer` | **High** | Pre-release gate: verifies `LicenseSalt.cs` is present and not the example, `ApplicationVersion` incremented, all new commands appear in `docs/catalog.md`, testing matrix updated |
| `test-generator` | **Medium** | Generates manual test cases for `docs/testing-matrix.md` from a `CommandDefinition` — covers happy path, empty selection, large range, undo, repeat, and license-locked scenarios |

Skip general-purpose agents (code reviewer, security scanner) for day-to-day work — the Hard Rules
and existing architecture make project-specific agents far more useful here.

---

## Claude Code Session Instructions

### Before every session

1. Read this file (you are doing so now).
2. Check `docs/architecture.md` for the full design rationale behind specific decisions.
3. **Do not open, read, or modify `LicenseSalt.cs` under any circumstances.**

### When adding a new tool

Follow the checklist above exactly. The registry is purely reflection-driven — no manual registration.
Every `[ExcelCommand]` class with a parameterless constructor and a valid `Definition.Id` is picked up.
The most common mistake is forgetting to add the `<Compile>` entry to `utilities.csproj`.

### When editing existing tools

- **Never change `CommandDefinition.Id`** — it is the ribbon control tag. Changing it silently
  breaks the ribbon for existing installations.
- **Never remove `UndoMode`** from a tool that had it — users rely on Ctrl+Z.
- `OperationRunner.RunGuarded` must remain the sole mutation path. Do not call Excel COM from
  `CommandBase.Execute()` before entering `RunGuarded`.

### When working on the ribbon

- `RibbonXmlBuilder` generates everything from `CommandDefinition`. Do not hardcode button XML.
- System controls (`sys.undo`, `sys.repeat`, `sys.findrun`, etc.) are defined as constants in
  `RibbonController`. Add new system controls there, not in `CommandRegistry`.
- To add a tab or group, update `TabRemap` or `GroupImageMso` in `RibbonXmlBuilder`.

### When working on the undo system

- Choose `UndoMode` based on expected cell count. `FullSnapshot` on a potentially large range is
  dangerous — prefer `PartialSnapshot` and call `ctx.Undo.RecordCell(cell)` before each write.
- `UndoMode.None` is correct for: read-only, navigation, selection, export, and tools that cannot
  meaningfully be reversed.
- The 200,000-cell limit for `FullSnapshot` is enforced in `UndoTransaction.CaptureUpFront`.
  Exceeding it silently skips undo capture (`Downgraded = true`).

### When working on services

- `ErrorService.Log` is the only sanctioned logging channel.
- `SettingsService` keys must be namespaced by command id: `"CommandId.SettingName"`. Use the
  typed accessors: `SettingsService.Get(key, defaultValue)`, `GetInt(key, default)`,
  `GetBool(key, default)`, and the matching `Set(key, value)` overloads. Read the live value in
  the dialog's OK handler rather than caching it at open time.
- `GridFocusService` owns the named shapes `__GF_Row__` and `__GF_Col__`. Never create shapes
  with those names elsewhere.
- `KeyboardHook` callback runs outside the COM-safe window. Only pure in-memory operations are
  permitted there.

### What NOT to do

- Do not add JSON/NuGet packages — this is .NET Framework 4.8 with no package manager.
- Do not use `async/await` in tool commands — Excel's STA COM model is incompatible with
  task-based asynchrony unless carefully marshalled to the UI thread.
- Do not call `Marshal.ReleaseComObject` — interop types are embedded; manual release causes
  double-free crashes.
- Do not hold `Excel.Range`, `Excel.Worksheet`, or `Excel.Workbook` as long-lived fields in services.
  Always re-resolve from the application context when needed.
- Do not copy patterns from `ExcelHelper.cs` or `DataConverterHelper.cs` — these are pre-framework
  legacy utilities that use `Globals.ThisAddIn.Application.Selection` directly and cell-by-cell
  iteration. All new tools use `ctx.Target` and bulk `Value2` reads/writes.
- **Do not open, read, or modify `LicenseSalt.cs`.**
