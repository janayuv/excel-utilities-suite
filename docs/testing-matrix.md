# Excel Utilities Suite — Testing Matrix

Run before every batch lands. Record PASS / FAIL(note) / N/A per tool per bucket.

---

## Test Buckets

| # | Bucket | Setup |
|---|---|---|
| T1 | Small range | 10×5 cells, mixed text/numbers |
| T2 | Large range | 50 000+ cells (column A, Sheet1) |
| T3 | Empty / no selection | Click outside used range, then run |
| T4 | Protected sheet | Review → Protect Sheet (no password) |
| T5 | Filtered / hidden rows | AutoFilter active, some rows hidden |
| T6 | Merged cells in range | 3 merged cells inside the target range |
| T7 | Multi-sheet scope | Tool set to Workbook scope |
| T8 | 32-bit Excel | Office 32-bit build |

## Expected Behaviour

| Bucket | Expected |
|---|---|
| T1 | Correct result; Ctrl+Z restores if UndoMode ≠ None |
| T2 | Perf toggles off; status-bar progress visible; undo respects UndoMode (FullSnapshot skipped with warning above 200k cells) |
| T3 | Guard message shown; no mutation; no crash |
| T4 | "Sheet is protected" message; no mutation; no crash |
| T5 | Visible-only vs all-cells behaviour matches tool documentation |
| T6 | Handled correctly OR refused with clear message; no corruption |
| T7 | All targeted sheets affected |
| T8 | Add-in loads; all tools run; no bitness COM errors |

---

## Framework Verification (run after any architecture change)

| Check | How | Result |
|---|---|---|
| Self-check PASS | Suite → Open Log after startup: expect `[Self-check] PASS — 113 commands` | |
| Screentip / supertip | Hover each button; bold title + description must appear | |
| Ctrl+Z undo | Change Case on 5 cells → Ctrl+Z → cells restored | |
| Ctrl+Z passthrough | Empty undo stack → Ctrl+Z → Excel native undo fires | |
| Error dialog | Force `throw` in any Run(); friendly dialog + log entry appear | |
| Status-bar progress | Run on 50k+ cells; `ExcelUtilities:` progress visible | |
| Excel state restored | After any error; ScreenUpdating/Calculation/EnableEvents back to original | |
| License stub | About dialog shows `License.Current.StatusText` | |

---

## Results Log

### Phase 2 — Parity (7 tools)

| Tool | T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 |
|---|---|---|---|---|---|---|---|---|
| Color Duplicates | | | | | | | | |
| Insert Sequence | | | | | | | | |
| Convert Text→Numbers | | | | | | | | |
| Convert Numbers→Text | | | | | | | | |
| Change Case | | | | | | | | |
| Trim / Clean Text | | | | | | | | |
| Export to PDF | | | | | | | | |

### Phase 3 — Top-10

| Tool | T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 |
|---|---|---|---|---|---|---|---|---|
| Add Text | | | | | | | | |
| Remove Characters | | | | | | | | |
| Remove Duplicate Rows | | | | | | | | |
| Select by Color | | | | | | | | |
| Sum/Count by Color | | | | | | | | |

### Phase 4 — Printing (12 tools)

| Tool | T1 | T2 | T3 | T4 | T5 | T6 | T7 | T8 |
|---|---|---|---|---|---|---|---|---|
| Print Multiple Workbooks | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print Multiple Selections | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print First Page Each Sheet | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print Pages Reverse Order | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print Current Page | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print Specified Pages | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print Circle Invalid Data | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Print Charts Only | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Copy Page Setup | N/A | N/A | N/A | N/A | N/A | N/A | N/A | |
| Paging Subtotals | | | N/A | | | N/A | N/A | |
| Insert Page Break Every Row | | | N/A | | | N/A | N/A | |
| Add Border to Each Page | | | N/A | | | N/A | N/A | |

---

*Cells: PASS · FAIL(description) · N/A · SKIP(reason)*

## Find & Run + Repeat Last Tool (v1.2)

| # | Scenario | Steps | Expected | Pass |
|---|---|---|---|---|
| FR1 | Repeat after direct tool | Run Proper Case on A1:A5; select B1:B5; click Repeat | B1:B5 proper-cased, no dialog | |
| FR2 | Repeat after dialog tool | Round Values to 2 dp on a range; select another range; click Repeat | Second range rounded to 2 dp, no dialog | |
| FR3 | Repeat, no selection | Repeat with nothing selectable | Friendly "select a range" warning, no crash | |
| FR4 | Repeat after excluded tool | Export to PDF; check Repeat label | Label unchanged (still prior repeatable tool) | |
| FR5 | Repeat with empty history | Fresh session; open Suite tab | Repeat button greyed/disabled | |
| FR6 | Find & Run filter | Open Find & Run; type "case" | List narrows to case tools | |
| FR7 | Find & Run keyboard | Type, press Down then Enter | Highlighted tool runs | |
| FR8 | Find & Run double-click | Double-click a tool | Tool runs | |
| FR9 | Find & Run launches dialog tool | Pick Add Text | Add Text dialog opens normally | |
| FR10 | Repeat reflects Find & Run | Run a tool via Find & Run; check Repeat label | "Repeat: <that tool>" | |
| FR11 | Startup self-check | Read log after load | "[Self-check] PASS — 113 commands registered, 0 definition warnings." | |
