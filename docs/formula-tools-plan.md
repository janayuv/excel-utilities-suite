# Formula Tools — Spell Number & Top 365 Formulas

**Branch:** feature/find-run-repeat-last  
**File:** `Utilities/Commands/Tools/FormulaCommands.cs` (new)  
**Tab:** Formula & Statistics → Utilities +  
**Group:** Formulas (existing)

---

## What's being added

### 1. Spell Number (multi-currency)

**Id:** `Formula.SpellNumber`  
**Label:** Spell Number  
**Order:** 40 (after WrapIferror @ 30, ToggleRefStyle @ 25)  
**UndoMode:** FullSnapshot  
**Icon:** `ImageMso = "NumberFormats"` (no custom PNG needed)

Converts numeric cell values to their spelled-out English equivalent.

| Value | Currency | Result |
|-------|----------|--------|
| 1234.56 | USD | One Thousand Two Hundred Thirty Four Dollars and 56 Cents |
| 500 | INR | Five Hundred Rupees |
| 99.01 | GBP | Ninety Nine Pounds and 1 Penny |
| 0 | EUR | Zero Euros |
| -42 | None | Negative Forty Two |

Supported currencies: USD, GBP, EUR, INR, AUD, CAD, SGD, JPY (no decimals), MYR, AED, None.  
Option: UPPERCASE output.  
Preview: shows conversion result for first numeric cell before applying.  
Settings: last-used currency persisted in `"Formula.SpellNumber.LastCurrency"`.

### 2. Insert Top 365 Formula

**Id:** `Formula.InsertTop365`  
**Label:** Top 365 Formulas  
**Order:** 41  
**UndoMode:** FullSnapshot  
**Icon:** `ImageMso = "InsertFunction"`

Searchable dialog listing 32 modern Excel 365 formulas across 7 categories:
Lookup, Dynamic Array, Text, Logic, Math, Lambda, Array.

Shows syntax reference + editable example — user edits, then inserts into active cell.

---

## csproj change

`<Compile Include="Commands\Tools\FormulaCommands.cs" />` added after DataAdvancedCommands.cs

---

## Verification checklist

- [ ] Build passes: `msbuild "Utilities\utilities.csproj" /p:Configuration=Debug /nologo /v:minimal`
- [ ] Self-check PASS in suite.log (N+2 commands, 0 warnings)
- [ ] Spell Number: numeric cells convert to text, undo restores
- [ ] Spell Number: non-numeric cells skipped
- [ ] Spell Number: last currency remembered across opens
- [ ] Top 365: search filters list; selecting entry shows syntax + description + example
- [ ] Top 365: Insert writes example to active cell, undo restores
- [ ] Both commands visible in Find & Run and in Utilities + → Formulas group
