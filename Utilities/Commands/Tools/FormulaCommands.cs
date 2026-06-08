using System;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using utilities.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace utilities.Commands.Tools
{
    // ══════════════════════════════════════════════════════════════════════════
    // Spell Number — convert numeric cell values to spelled-out currency text
    // ══════════════════════════════════════════════════════════════════════════

    [ExcelCommand]
    public sealed class SpellNumberCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id       = "Formula.SpellNumber",
            Label    = "Spell Number",
            Screentip = "Spell Number to Words",
            Supertip = "Convert numeric values in selected cells to spelled-out text with optional currency — e.g. 1 234.56 → \"One Thousand Two Hundred Thirty Four Dollars and 56 Cents\".",
            ImageMso = "NumberFormats",
            Tab      = "Formula & Statistics",
            Group    = "Formulas",
            Order    = 40,
            Scope    = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot,
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new SpellNumberDialog();
    }

    internal sealed class SpellNumberDialog : Dialogs.DialogBase
    {
        // ── Currency catalogue ────────────────────────────────────────────────
        internal struct CurrencyDef
        {
            public readonly string DisplayName;
            public readonly string SingularMain;
            public readonly string PluralMain;
            public readonly bool   HasDecimal;
            public readonly string SingularDec;
            public readonly string PluralDec;

            public CurrencyDef(string name, string singM, string plurM,
                               bool hasDec, string singD, string plurD)
            {
                DisplayName  = name;
                SingularMain = singM; PluralMain = plurM;
                HasDecimal   = hasDec;
                SingularDec  = singD; PluralDec  = plurD;
            }
        }

        private static readonly CurrencyDef[] Currencies =
        {
            new CurrencyDef("USD - US Dollar",           "Dollar",  "Dollars",  true,  "Cent",  "Cents" ),
            new CurrencyDef("GBP - British Pound",       "Pound",   "Pounds",   true,  "Penny", "Pence" ),
            new CurrencyDef("EUR - Euro",                "Euro",    "Euros",    true,  "Cent",  "Cents" ),
            new CurrencyDef("INR - Indian Rupee",        "Rupee",   "Rupees",   true,  "Paisa", "Paise" ),
            new CurrencyDef("AUD - Australian Dollar",   "Dollar",  "Dollars",  true,  "Cent",  "Cents" ),
            new CurrencyDef("CAD - Canadian Dollar",     "Dollar",  "Dollars",  true,  "Cent",  "Cents" ),
            new CurrencyDef("SGD - Singapore Dollar",    "Dollar",  "Dollars",  true,  "Cent",  "Cents" ),
            new CurrencyDef("JPY - Japanese Yen",        "Yen",     "Yen",      false, null,    null    ),
            new CurrencyDef("MYR - Malaysian Ringgit",   "Ringgit", "Ringgit",  true,  "Sen",   "Sen"   ),
            new CurrencyDef("AED - UAE Dirham",          "Dirham",  "Dirhams",  true,  "Fil",   "Fils"  ),
            new CurrencyDef("(None - numbers only)",     null,      null,       false, null,    null    ),
        };

        // ── Controls ──────────────────────────────────────────────────────────
        private readonly ComboBox _currencyCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly CheckBox _upperCase      = new CheckBox { Text = "UPPERCASE output", AutoSize = true };
        private readonly Label    _preview        = new Label   { AutoSize = false, BorderStyle = BorderStyle.FixedSingle, TextAlign = ContentAlignment.MiddleLeft };

        public SpellNumberDialog()
        {
            Text       = SpellNumberCommand.Def.Label;
            ClientSize = new Size(420, 188);

            var lblCur = new Label { Text = "Currency:", Left = 12, Top = 17, AutoSize = true };
            _currencyCombo.SetBounds(88, 13, 316, 23);
            foreach (CurrencyDef c in Currencies) _currencyCombo.Items.Add(c.DisplayName);

            string lastCur = SettingsService.Get("Formula.SpellNumber.LastCurrency", Currencies[0].DisplayName);
            int savedIdx = _currencyCombo.Items.IndexOf(lastCur);
            _currencyCombo.SelectedIndex = savedIdx >= 0 ? savedIdx : 0;
            _currencyCombo.SelectedIndexChanged += (s, e) => UpdatePreview();

            _upperCase.SetBounds(88, 44, 200, 22);
            _upperCase.CheckedChanged += (s, e) => UpdatePreview();

            var lblPrev = new Label { Text = "Preview:", Left = 12, Top = 78, AutoSize = true };
            _preview.SetBounds(12, 97, 392, 44);

            var btnApply  = new Button { Text = "&Apply",  Left = 224, Top = 152, Width = 84 };
            var btnCancel = new Button { Text = "&Cancel", Left = 314, Top = 152, Width = 84, DialogResult = DialogResult.Cancel };
            btnApply.Click += OnApply;

            Controls.AddRange(new Control[] { lblCur, _currencyCombo, _upperCase, lblPrev, _preview, btnApply, btnCancel });
            WireButtons(btnApply, btnCancel);

            Shown += (s, e) => UpdatePreview();
        }

        private CurrencyDef SelectedCurrency => Currencies[_currencyCombo.SelectedIndex];

        private void UpdatePreview()
        {
            try
            {
                Excel.Range sel = CurrentSelection;
                if (sel == null) { _preview.Text = "(no selection)"; return; }

                foreach (Excel.Range cell in sel.Cells)
                {
                    if (!(cell.Value2 is double)) continue;
                    decimal d;
                    try { d = Convert.ToDecimal((double)cell.Value2); }
                    catch { continue; }
                    string spelled = SpellAmount(d, SelectedCurrency);
                    _preview.Text = _upperCase.Checked ? spelled.ToUpper() : spelled;
                    return;
                }
                _preview.Text = "(select numeric cells to preview)";
            }
            catch { _preview.Text = ""; }
        }

        private void OnApply(object sender, EventArgs e)
        {
            bool upper = _upperCase.Checked;
            CurrencyDef cur = SelectedCurrency;
            SettingsService.Set("Formula.SpellNumber.LastCurrency", cur.DisplayName);

            bool ok = RunOperation(SpellNumberCommand.Def, CurrentSelection, ctx =>
            {
                int total = ctx.Target.Cells.Count;
                int done  = 0;

                foreach (Excel.Range area in ctx.Target.Areas)
                {
                    object[,] values = area.Value2 as object[,];
                    if (values == null)
                    {
                        if (area.Value2 is double d)
                        {
                            string s = TrySpell(d, cur, upper);
                            if (s != null) area.Value2 = s;
                        }
                        done++;
                        ctx.Progress.Report((double)done / total);
                        continue;
                    }

                    int rows = values.GetLength(0), cols = values.GetLength(1);
                    for (int r = 1; r <= rows; r++)
                        for (int c = 1; c <= cols; c++)
                        {
                            if (values[r, c] is double dv)
                            {
                                string s = TrySpell(dv, cur, upper);
                                if (s != null) values[r, c] = s;
                            }
                            done++;
                            if ((done & 0xFF) == 0) ctx.Progress.Report((double)done / total);
                        }
                    area.Value2 = values;
                }
            });

            if (ok) Close();
        }

        private static string TrySpell(double d, CurrencyDef cur, bool upper)
        {
            try
            {
                decimal dec = Convert.ToDecimal(d);
                string s = SpellAmount(dec, cur);
                return upper ? s.ToUpper() : s;
            }
            catch { return null; }
        }

        // ── Number-to-words engine ────────────────────────────────────────────

        private static readonly string[] _ones =
        {
            "Zero","One","Two","Three","Four","Five","Six","Seven","Eight","Nine",
            "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen",
            "Seventeen","Eighteen","Nineteen"
        };
        private static readonly string[] _tens =
            { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

        private static string NumberToWords(long n)
        {
            if (n == 0) return "Zero";
            if (n < 0)  return "Negative " + NumberToWords(-n);

            var sb = new StringBuilder();
            if (n >= 1000000000000L) { sb.Append(NumberToWords(n / 1000000000000L)).Append(" Trillion "); n %= 1000000000000L; }
            if (n >= 1000000000L)    { sb.Append(NumberToWords(n / 1000000000L)).Append(" Billion ");  n %= 1000000000L; }
            if (n >= 1000000L)       { sb.Append(NumberToWords(n / 1000000L)).Append(" Million ");     n %= 1000000L; }
            if (n >= 1000L)          { sb.Append(NumberToWords(n / 1000L)).Append(" Thousand ");       n %= 1000L; }
            if (n >= 100L)           { sb.Append(_ones[n / 100]).Append(" Hundred ");                  n %= 100L; }
            if (n >= 20L)
            {
                sb.Append(_tens[n / 10]);
                if (n % 10 > 0) sb.Append(" ").Append(_ones[n % 10]);
            }
            else if (n > 0L)
                sb.Append(_ones[n]);

            return sb.ToString().Trim();
        }

        internal static string SpellAmount(decimal amount, CurrencyDef cur)
        {
            bool negative = amount < 0;
            decimal abs   = Math.Abs(amount);
            long intPart  = (long)Math.Truncate(abs);
            int  decPart  = cur.HasDecimal ? (int)Math.Round((abs - intPart) * 100) : 0;

            var sb = new StringBuilder();
            if (negative) sb.Append("Negative ");

            if (intPart == 0 && decPart == 0)
            {
                sb.Append("Zero");
                if (cur.PluralMain != null) sb.Append(" ").Append(cur.PluralMain);
                return sb.ToString();
            }

            if (intPart > 0)
            {
                sb.Append(NumberToWords(intPart));
                if (cur.SingularMain != null)
                    sb.Append(" ").Append(intPart == 1 ? cur.SingularMain : cur.PluralMain);
            }

            if (cur.HasDecimal && decPart > 0)
            {
                if (intPart > 0) sb.Append(" and ");
                sb.Append(NumberToWords(decPart));
                sb.Append(" ").Append(decPart == 1 ? cur.SingularDec : cur.PluralDec);
            }

            return sb.ToString();
        }
    }

    // ══════════════════════════════════════════════════════════════════════════
    // Insert Top 365 Formula — browse and insert modern Excel 365 formula templates
    // ══════════════════════════════════════════════════════════════════════════

    [ExcelCommand]
    public sealed class InsertTopFormulaCommand : DialogCommandBase
    {
        public static readonly CommandDefinition Def = new CommandDefinition
        {
            Id       = "Formula.InsertTop365",
            Label    = "Top 365 Formulas",
            Screentip = "Insert Top Office 365 Formula",
            Supertip = "Browse and insert modern Excel 365 formula templates - XLOOKUP, FILTER, UNIQUE, LAMBDA, LET, TEXTSPLIT and more.",
            ImageMso = "InsertFunction",
            Tab      = "Formula & Statistics",
            Group    = "Formulas",
            Order    = 41,
            Scope    = CommandScope.Selection,
            UndoMode = UndoMode.FullSnapshot,
        };

        public override CommandDefinition Definition => Def;
        protected override Dialogs.DialogBase CreateDialog() => new InsertTopFormulaDialog();
    }

    internal sealed class InsertTopFormulaDialog : Dialogs.DialogBase
    {
        private struct FormulaEntry
        {
            public readonly string Category;
            public readonly string Name;
            public readonly string Description;
            public readonly string Syntax;
            public readonly string Example;

            public FormulaEntry(string cat, string name, string desc, string syntax, string example)
            {
                Category = cat; Name = name; Description = desc; Syntax = syntax; Example = example;
            }
        }

        private static readonly FormulaEntry[] AllFormulas =
        {
            // Lookup & Reference
            new FormulaEntry("Lookup", "XLOOKUP",
                "Searches any direction and returns matched values. Replaces VLOOKUP, HLOOKUP, and INDEX/MATCH. Returns the exact match by default.",
                "=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])",
                "=XLOOKUP(A2, B:B, C:C, \"Not found\")"),
            new FormulaEntry("Lookup", "XMATCH",
                "Returns the relative position of a lookup value in a range or array. More flexible than MATCH.",
                "=XMATCH(lookup_value, lookup_array, [match_mode], [search_mode])",
                "=XMATCH(\"Alice\", A2:A20)"),
            // Dynamic Arrays
            new FormulaEntry("Dynamic Array", "FILTER",
                "Returns rows or columns that meet one or more conditions as a spill array. Combine with SORT or UNIQUE for powerful reporting.",
                "=FILTER(array, include, [if_empty])",
                "=FILTER(A2:C20, B2:B20>100, \"No results\")"),
            new FormulaEntry("Dynamic Array", "UNIQUE",
                "Returns unique distinct values from a range or array. Set exactly_once=TRUE to return values that appear only once.",
                "=UNIQUE(array, [by_col], [exactly_once])",
                "=UNIQUE(A2:A100)"),
            new FormulaEntry("Dynamic Array", "SORT",
                "Sorts the contents of a range or array. sort_order: 1=ascending (default), -1=descending.",
                "=SORT(array, [sort_index], [sort_order], [by_col])",
                "=SORT(A2:B20, 2, -1)"),
            new FormulaEntry("Dynamic Array", "SORTBY",
                "Sorts a range or array based on corresponding values in one or more other arrays.",
                "=SORTBY(array, by_array1, [sort_order1], [by_array2, sort_order2], ...)",
                "=SORTBY(A2:C10, C2:C10, -1)"),
            new FormulaEntry("Dynamic Array", "SEQUENCE",
                "Generates a sequence of numbers in a spill array. Useful for row numbers, calendars, and date series.",
                "=SEQUENCE(rows, [cols], [start], [step])",
                "=SEQUENCE(10, 1, 1, 2)"),
            new FormulaEntry("Dynamic Array", "RANDARRAY",
                "Returns an array of random numbers. Set whole_number=TRUE for integers.",
                "=RANDARRAY([rows], [cols], [min], [max], [whole_number])",
                "=RANDARRAY(5, 3, 1, 100, TRUE)"),
            // Text
            new FormulaEntry("Text", "TEXTJOIN",
                "Joins multiple text strings with a delimiter. Set ignore_empty=TRUE to skip blank cells.",
                "=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)",
                "=TEXTJOIN(\", \", TRUE, A2:A10)"),
            new FormulaEntry("Text", "TEXTSPLIT",
                "Splits a text string into rows and columns using delimiters. Returns a spill array.",
                "=TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with])",
                "=TEXTSPLIT(A1, \",\")"),
            new FormulaEntry("Text", "TEXTBEFORE",
                "Returns text that occurs before a given delimiter. Use instance_num=-1 for the last occurrence.",
                "=TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])",
                "=TEXTBEFORE(A1, \" \")"),
            new FormulaEntry("Text", "TEXTAFTER",
                "Returns text that occurs after a given delimiter. Complements TEXTBEFORE.",
                "=TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])",
                "=TEXTAFTER(A1, \"-\")"),
            new FormulaEntry("Text", "CONCAT",
                "Concatenates values from multiple cells or ranges without a separator (use TEXTJOIN to add one).",
                "=CONCAT(text1, [text2], ...)",
                "=CONCAT(A2:A10)"),
            // Logic
            new FormulaEntry("Logic", "IFS",
                "Tests multiple conditions in order and returns the first matching result. Add TRUE as the last condition for a default value.",
                "=IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)",
                "=IFS(A1>90,\"A\", A1>80,\"B\", A1>70,\"C\", TRUE,\"F\")"),
            new FormulaEntry("Logic", "SWITCH",
                "Compares an expression against a list of values and returns the matching result. Cleaner than nested IFS for equality checks.",
                "=SWITCH(expression, value1, result1, [value2, result2], ..., [default])",
                "=SWITCH(WEEKDAY(A1),1,\"Sun\",2,\"Mon\",3,\"Tue\",4,\"Wed\",5,\"Thu\",6,\"Fri\",\"Sat\")"),
            // Math & Stats
            new FormulaEntry("Math", "MAXIFS",
                "Returns the maximum value from cells that meet one or more criteria. Equivalent to a conditional MAX.",
                "=MAXIFS(max_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
                "=MAXIFS(C2:C100, B2:B100, \"Sales\")"),
            new FormulaEntry("Math", "MINIFS",
                "Returns the minimum value from cells that meet one or more criteria. Equivalent to a conditional MIN.",
                "=MINIFS(min_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)",
                "=MINIFS(C2:C100, A2:A100, \"East\")"),
            // Lambda & Advanced
            new FormulaEntry("Lambda", "LET",
                "Assigns names to intermediate calculation results for reuse within the formula. Improves readability and avoids repeated calculations.",
                "=LET(name1, value1, [name2, value2], ..., calculation)",
                "=LET(filtered, FILTER(A2:A100, B2:B100>50), SUM(filtered))"),
            new FormulaEntry("Lambda", "LAMBDA",
                "Defines a custom reusable function inline. Use with Name Manager to create a named function callable anywhere in the workbook.",
                "=LAMBDA([parameter1], [parameter2], ..., body)",
                "=LAMBDA(x, y, SQRT(x*x + y*y))(3, 4)"),
            new FormulaEntry("Lambda", "MAP",
                "Creates a new array by applying a LAMBDA function to each element of one or more arrays.",
                "=MAP(array1, [array2], ..., lambda)",
                "=MAP(A2:A10, LAMBDA(x, x*x))"),
            new FormulaEntry("Lambda", "REDUCE",
                "Reduces an array to a single accumulated value by applying a LAMBDA function cumulatively.",
                "=REDUCE([initial_value], array, lambda)",
                "=REDUCE(0, A2:A10, LAMBDA(acc, val, acc+val))"),
            new FormulaEntry("Lambda", "SCAN",
                "Like REDUCE but returns an array of all intermediate accumulated values (running totals, running counts).",
                "=SCAN([initial_value], array, lambda)",
                "=SCAN(0, A2:A10, LAMBDA(acc, val, acc+val))"),
            new FormulaEntry("Lambda", "BYROW",
                "Applies a LAMBDA to each row of an array and returns an array of per-row results.",
                "=BYROW(array, lambda)",
                "=BYROW(A2:C10, LAMBDA(row, SUM(row)))"),
            new FormulaEntry("Lambda", "BYCOL",
                "Applies a LAMBDA to each column of an array and returns an array of per-column results.",
                "=BYCOL(array, lambda)",
                "=BYCOL(A2:C10, LAMBDA(col, AVERAGE(col)))"),
            // Array Manipulation
            new FormulaEntry("Array", "VSTACK",
                "Stacks two or more arrays vertically (appending rows) into a single array.",
                "=VSTACK(array1, [array2], ...)",
                "=VSTACK(A2:B5, D2:E8)"),
            new FormulaEntry("Array", "HSTACK",
                "Appends two or more arrays horizontally (appending columns) into a single array.",
                "=HSTACK(array1, [array2], ...)",
                "=HSTACK(A2:A10, C2:C10, E2:E10)"),
            new FormulaEntry("Array", "TOCOL",
                "Transforms a 2D array into a single column vector. Use ignore=1 to drop blanks.",
                "=TOCOL(array, [ignore], [scan_by_column])",
                "=TOCOL(A2:D10)"),
            new FormulaEntry("Array", "TOROW",
                "Transforms a 2D array into a single row vector.",
                "=TOROW(array, [ignore], [scan_by_column])",
                "=TOROW(A2:D10)"),
            new FormulaEntry("Array", "CHOOSECOLS",
                "Returns only the specified columns from an array by column index.",
                "=CHOOSECOLS(array, col_num1, [col_num2], ...)",
                "=CHOOSECOLS(A2:E10, 1, 3, 5)"),
            new FormulaEntry("Array", "CHOOSEROWS",
                "Returns only the specified rows from an array by row index.",
                "=CHOOSEROWS(array, row_num1, [row_num2], ...)",
                "=CHOOSEROWS(A2:E10, 1, 3, 5)"),
            new FormulaEntry("Array", "WRAPROWS",
                "Wraps a 1D vector into a 2D array by filling row by row.",
                "=WRAPROWS(vector, wrap_count, [pad_with])",
                "=WRAPROWS(A2:A20, 4)"),
            new FormulaEntry("Array", "WRAPCOLS",
                "Wraps a 1D vector into a 2D array by filling column by column.",
                "=WRAPCOLS(vector, wrap_count, [pad_with])",
                "=WRAPCOLS(A2:A20, 4)"),
            new FormulaEntry("Array", "EXPAND",
                "Expands or pads an array to the specified number of rows and columns.",
                "=EXPAND(array, rows, [columns], [pad_with])",
                "=EXPAND(A2:B3, 5, 3, 0)"),
        };

        // ── Controls ──────────────────────────────────────────────────────────
        private readonly TextBox _search     = new TextBox();
        private readonly ListBox _list       = new ListBox { IntegralHeight = false };
        private readonly Label   _lblName    = new Label   { AutoSize = true };
        private readonly Label   _lblCat     = new Label   { AutoSize = true };
        private readonly TextBox _descBox    = new TextBox { ReadOnly = true, Multiline = true, ScrollBars = ScrollBars.Vertical };
        private readonly Label   _lblSyntax  = new Label   { Text = "Syntax:", AutoSize = true };
        private readonly TextBox _syntaxBox  = new TextBox { ReadOnly = true, Multiline = true };
        private readonly Label   _lblExample = new Label   { Text = "Example to insert (edit as needed):", AutoSize = true };
        private readonly TextBox _exampleBox = new TextBox { Multiline = true };

        private FormulaEntry[] _filtered = AllFormulas;

        public InsertTopFormulaDialog()
        {
            Text       = InsertTopFormulaCommand.Def.Label;
            ClientSize = new Size(680, 468);

            var lblSearch = new Label { Text = "Search:", Left = 12, Top = 16, AutoSize = true };
            _search.SetBounds(68, 12, 596, 23);
            _search.TextChanged += (s, e) => RefreshList();

            _list.SetBounds(12, 44, 198, 380);
            _list.SelectedIndexChanged += (s, e) => ShowSelected();

            _lblName.SetBounds(222, 44, 444, 20);
            _lblName.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
            _lblCat.SetBounds(222, 66, 444, 16);
            _lblCat.ForeColor = Color.Gray;

            _descBox.SetBounds(222, 88, 444, 78);
            _descBox.BackColor = Color.FromArgb(248, 248, 248);

            _lblSyntax.SetBounds(222, 174, 444, 16);
            _syntaxBox.SetBounds(222, 192, 444, 46);
            _syntaxBox.Font = new Font("Courier New", 8f);
            _syntaxBox.BackColor = Color.FromArgb(240, 240, 240);

            _lblExample.SetBounds(222, 246, 444, 16);
            _exampleBox.SetBounds(222, 264, 444, 32);
            _exampleBox.Font = new Font("Courier New", 8f);

            var note = new Label
            {
                Text      = "Requires Excel 365 subscription. Edit the example above then click Insert.",
                Left      = 222, Top = 304, Width = 444, Height = 28,
                ForeColor = Color.Gray, AutoSize = false
            };

            var btnInsert = new Button { Text = "&Insert Formula", Left = 480, Top = 428, Width = 100 };
            var btnCancel = new Button { Text = "&Cancel",        Left = 584, Top = 428, Width = 84, DialogResult = DialogResult.Cancel };
            btnInsert.Click += OnInsert;

            Controls.AddRange(new Control[]
            {
                lblSearch, _search,
                _list,
                _lblName, _lblCat, _descBox,
                _lblSyntax, _syntaxBox,
                _lblExample, _exampleBox, note,
                btnInsert, btnCancel,
            });
            WireButtons(btnInsert, btnCancel);

            RefreshList();
        }

        private void RefreshList()
        {
            string q = _search.Text.Trim();
            _filtered = string.IsNullOrEmpty(q)
                ? AllFormulas
                : AllFormulas.Where(f =>
                    f.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    f.Category.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0 ||
                    f.Description.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0
                ).ToArray();

            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (FormulaEntry f in _filtered) _list.Items.Add(f.Name);
            _list.EndUpdate();

            if (_list.Items.Count > 0)
                _list.SelectedIndex = 0;
            else
                ClearDetail();
        }

        private void ShowSelected()
        {
            int i = _list.SelectedIndex;
            if (i < 0 || i >= _filtered.Length) { ClearDetail(); return; }
            FormulaEntry f = _filtered[i];
            _lblName.Text    = f.Name;
            _lblCat.Text     = f.Category;
            _descBox.Text    = f.Description;
            _syntaxBox.Text  = f.Syntax;
            _exampleBox.Text = f.Example;
        }

        private void ClearDetail()
        {
            _lblName.Text = _lblCat.Text = _descBox.Text = _syntaxBox.Text = _exampleBox.Text = "";
        }

        private void OnInsert(object sender, EventArgs e)
        {
            if (_list.SelectedIndex < 0)
            {
                MessageBox.Show("Please select a formula from the list.",
                    InsertTopFormulaCommand.Def.Label, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string template = _exampleBox.Text.Trim();
            if (string.IsNullOrEmpty(template)) return;

            bool ok = RunOperation(InsertTopFormulaCommand.Def, CurrentSelection, ctx =>
            {
                Excel.Range target = (Excel.Range)ctx.Target.Cells[1, 1];
                if (template.StartsWith("="))
                    target.Formula = template;
                else
                    target.Value2 = template;
            });

            if (ok) Close();
        }
    }
}
