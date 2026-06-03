# Excel Utilities Suite — Command Catalog

> Derived from `CommandDefinition` metadata across all `[ExcelCommand]` classes.
> **113 commands** · 2 ribbon tabs · 15 groups.
> Runtime tab remapping: see `RibbonXmlBuilder.TabRemap`.

---

## Ribbon: Utilities

### View

| # | Id | Label | Undo |
|---|---|---|---|
| 5 | View.GridFocus | Grid Focus | None |

### Range & Cells

| # | Id | Label | Undo |
|---|---|---|---|
| 20 | Range.ColorDuplicates | Color Duplicates | None |
| 30 | Range.DeleteBlankRows | Delete Blank Rows | None |
| 31 | Range.InsertBlankRows | Insert Blank Rows | None |
| 40 | Range.FillDownBlanks | Fill Down Blanks | FullSnapshot |
| 50 | Range.Transpose | Transpose Range | FullSnapshot |
| 55 | Range.FlipVertical | Flip Vertical | FullSnapshot |
| 56 | Range.FlipHorizontal | Flip Horizontal | FullSnapshot |
| 60 | Range.CompareRanges | Compare Ranges | None |
| 70 | Range.SwapRanges | Swap Ranges | FullSnapshot |
| 80 | Range.CopyVisibleCells | Copy Visible Cells | None |

---

## Ribbon: Utilities +

### Text

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Text.TrimText | Trim / Clean Text | None |
| 20 | Text.AddText | Add Text | FullSnapshot |
| 21 | Text.RemoveCharacters | Remove Characters | FullSnapshot |
| 30 | Text.Reverse | Reverse Text | FullSnapshot |
| 31 | Text.ExtractNumbers | Extract Numbers | FullSnapshot |
| 32 | Text.ExtractText | Extract Text | FullSnapshot |
| 33 | Text.SplitNames | Split Names | None |
| 50 | Text.FindReplaceAcrossSheets | Find & Replace All Sheets | None |
| 60 | Text.CountWords | Count Words | None |
| 61 | Text.AddLeadingZeros | Add Leading Zeros | FullSnapshot |
| 62 | Text.RemoveApostrophes | Remove Apostrophes | FullSnapshot |
| 63 | Text.ProperCase | Proper Case | FullSnapshot |
| 64 | Text.AddLineBreak | Add Line Break | FullSnapshot |
| 65 | Text.SuperSubscript | Super/Subscript | FullSnapshot |

### Format & Convert

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Text.ChangeCase | Change Case | None |
| 10 | Convert.TextToNumbers | Text to Numbers | FullSnapshot |
| 11 | Convert.NumbersToText | Numbers to Text | FullSnapshot |
| 15 | Convert.NumbersToWords | Numbers to Words | FullSnapshot |
| 20 | Format.RoundValues | Round Values | FullSnapshot |
| 21 | Format.ChangeSign | Change Sign | FullSnapshot |
| 25 | Convert.DateFormat | Convert Date Format | None |
| 30 | Format.ClearFormatting | Clear Formatting | None |
| 35 | Format.CopyCellFormatting | Copy Cell Formatting | None |
| 36 | Format.AlternateRows | Alternate Row Colors | None |
| 37 | Format.Currency | Currency Format | None |
| 40 | Convert.Units | Unit Conversion | FullSnapshot |
| 50 | Format.ClearHyperlinks | Clear Hyperlinks | None |
| 52 | Format.ClearConditionalFormatting | Clear Cond. Formatting | None |
| 55 | Format.AutoFitAll | AutoFit All | None |

### Insert

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Insert.SequenceNumbers | Insert Sequence | None |
| 20 | Insert.RandomData | Insert Random Data | FullSnapshot |
| 30 | Insert.Bullets | Insert Bullets | FullSnapshot |
| 40 | Insert.DateSequence | Insert Date Sequence | FullSnapshot |

### Select & Navigate

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Select.ByColor | Select by Color | None |
| 11 | Select.ByFontColor | Select by Font Color | None |
| 20 | Select.BlankCells | Select Blank Cells | None |
| 21 | Select.NonBlankCells | Select Non-Blank Cells | None |
| 22 | Select.ErrorCells | Select Error Cells | None |
| 30 | Select.MaxCell | Select Max Cell | None |
| 31 | Select.MinCell | Select Min Cell | None |
| 40 | Select.DuplicateCells | Select Duplicates | None |
| 41 | Select.UniqueCells | Select Uniques | None |
| 50 | Select.ByValue | Select by Value | None |
| 60 | Select.FirstCell | Select First Cell | None |
| 61 | Select.LastCell | Select Last Cell | None |

### Formula & Statistics

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Formula.SumCountByColor | Sum/Count by Color | None |
| 15 | Formula.CountByColor | Count by Color | None |
| 20 | Formula.CalculateAge | Calculate Age | None |
| 20 | Formula.ReplaceWithValues | Formulas to Values | FormulaOnly |
| 25 | Formula.ToggleRefStyle | Toggle Ref Style | FormulaOnly |
| 30 | Formula.WrapIferror | Wrap with IFERROR | FormulaOnly |
| 10 | Formula.CombineRows | Combine Rows | FullSnapshot |

### Data & Cleaning

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Data.RemoveDuplicateRows | Remove Duplicate Rows | None |
| 15 | Data.HighlightDuplicates | Highlight Duplicates | None |
| 20 | Data.FuzzyDedupe | Fuzzy Dedupe | None |
| 10 | Data.MergeCellsKeepData | Merge Cells (Keep Data) | FullSnapshot |
| 20 | Data.UnmergeAndFill | Unmerge and Fill | FullSnapshot |
| 10 | Data.SplitIntoRows | Split into Rows | None |
| 30 | Data.DetectTypes | Detect Data Types | None |

### Export / Import

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Export.ToPdf | Export to PDF | None |
| 20 | Export.SheetsToFiles | Sheets to Files | None |
| 30 | Export.ToCsv | Export to CSV | None |
| 31 | Export.ToJson | Export to JSON | None |
| 40 | Export.RangeAsImage | Range as Image | None |
| 10 | Import.FileList | Insert File List | None |
| 20 | Import.Filenames | Import Filenames | None |

### Workbook & Sheets

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Sheet.RenameMultiple | Rename Sheets | None |
| 15 | Sheet.BatchRename | Batch Rename Sheets | None |
| 20 | Sheet.Sort | Sort Sheets | None |
| 30 | Sheet.CreateTOC | Sheet TOC | None |
| 40 | Sheet.CopySheets | Copy Sheets | None |
| 50 | Sheet.UnhideAll | Unhide All Sheets | None |
| 10 | Sheet.MergeWorkbooks | Merge Workbooks | None |
| 20 | Sheet.SplitByColumn | Split Sheet by Column | None |
| 30 | Sheet.RefreshPivots | Refresh All Pivots | None |
| 10 | Format.FreezePanes | Freeze Panes | None |

### Printing

| # | Id | Label | Undo |
|---|---|---|---|
| 10 | Printing.PrintMultipleWorkbooks | Print Multiple Workbooks Wizard... | None |
| 20 | Printing.PrintMultipleSelections | Print Multiple Selections Wizard... | None |
| 30 | Printing.PrintFirstPage | Print First Page of Each Worksheet | None |
| 40 | Printing.PrintReverseOrder | Print Pages in Reverse Order | None |
| 50 | Printing.PrintCurrentPage | Print Current Page | None |
| 60 | Printing.PrintSpecifiedPages | Print Specified Pages... | None |
| 70 | Printing.PrintCircleInvalidData | Print Circle Invalid Data... | None |
| 80 | Printing.PrintChartsOnly | Print Charts Only... | None |
| 90 | Printing.CopyPageSetup | Copy Page Setup... | None |
| 100 | Printing.PagingSubtotals | Paging Subtotals... | None |
| 110 | Printing.InsertPageBreakEveryRow | Insert Page Break Every Row... | None |
| 120 | Printing.AddBorderToEachPage | Add Border to Each Page | None |

---

## Suite Tab (system controls)

| Id | Label | Notes |
|---|---|---|
| sys.undo | Undo Last Action | Dynamic label; enabled when stack non-empty |
| sys.about | About | Version + license status |
| sys.openlog | Open Log | Opens %APPDATA% error log |

---

**Total: 113 tool commands + 3 system controls = 116 ribbon buttons**

*UndoMode breakdown — None: 79 · FullSnapshot: 25 · FormulaOnly: 3 · PartialSnapshot: 0*
