using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using utilities.Commands;

namespace utilities.Dialogs
{
    /// <summary>
    /// A searchable picker over every registered command. Type to filter by label, id or
    /// tooltip; Up/Down moves the highlight; Enter / double-click chooses a tool. The chosen
    /// command id is exposed via <see cref="SelectedCommandId"/> after the dialog closes with
    /// OK; the caller (RibbonController) dispatches it so a chosen dialog tool opens cleanly.
    /// </summary>
    internal sealed class FindRunDialog : DialogBase
    {
        private readonly TextBox _search = new TextBox();
        private readonly ListView _list = new ListView();
        private readonly List<CommandDefinition> _all;

        /// <summary>The command id the user chose to run, or null if cancelled.</summary>
        public string SelectedCommandId { get; private set; }

        public FindRunDialog()
        {
            Text = "Find & Run a Utility";
            ClientSize = new Size(480, 420);
            MinimumSize = new Size(420, 320);
            FormBorderStyle = FormBorderStyle.Sizable; // overrides DialogBase's FixedDialog
            SizeGripStyle = SizeGripStyle.Show;

            _search.SetBounds(12, 12, 456, 23);
            _search.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            _search.TextChanged += (s, e) => RefreshList();
            _search.KeyDown += OnSearchKeyDown;

            _list.SetBounds(12, 44, 456, 364);
            _list.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _list.View = View.Details;
            _list.FullRowSelect = true;
            _list.MultiSelect = false;
            _list.HideSelection = false;
            _list.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            _list.Columns.Add("Tool", 170);
            _list.Columns.Add("Group", 110);
            _list.Columns.Add("Description", 170);
            _list.DoubleClick += (s, e) => RunSelected();

            Controls.Add(_search);
            Controls.Add(_list);

            // Snapshot the registry once, sorted for stable display.
            _all = CommandRegistry.All
                .Select(c => c.Definition)
                .Where(d => d != null && !string.IsNullOrEmpty(d.Label))
                .OrderBy(d => d.Tab, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.Group, StringComparer.OrdinalIgnoreCase)
                .ThenBy(d => d.Order)
                .ToList();

            RefreshList();
            ActiveControl = _search;
        }

        private void RefreshList()
        {
            string q = _search.Text.Trim();
            IEnumerable<CommandDefinition> matches = string.IsNullOrEmpty(q)
                ? _all
                : _all.Where(d => Matches(d, q));

            // Pin the last-used tool to the top when not actively searching.
            if (string.IsNullOrEmpty(q) && utilities.Services.RepeatService.CanRepeat)
            {
                string lastId = utilities.Services.RepeatService.LastId;
                matches = matches
                    .OrderByDescending(d => string.Equals(d.Id, lastId, StringComparison.Ordinal))
                    .ToList();
            }

            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (CommandDefinition d in matches)
            {
                var item = new ListViewItem(d.Label) { Tag = d.Id };
                item.SubItems.Add(d.Group ?? string.Empty);
                item.SubItems.Add(d.Supertip ?? d.Screentip ?? string.Empty);
                _list.Items.Add(item);
            }
            if (_list.Items.Count > 0)
            {
                _list.Items[0].Selected = true;
                _list.Items[0].Focused = true;
            }
            _list.EndUpdate();
        }

        /// <summary>Pure match predicate: case-insensitive substring over label/id/tooltip.</summary>
        private static bool Matches(CommandDefinition d, string query)
        {
            string ci = query.ToLowerInvariant();
            if (d.Label != null && d.Label.ToLowerInvariant().Contains(ci)) return true;
            if (d.Id != null && d.Id.ToLowerInvariant().Contains(ci)) return true;
            if (d.Supertip != null && d.Supertip.ToLowerInvariant().Contains(ci)) return true;
            return false;
        }

        private void OnSearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                MoveSelection(e.KeyCode == Keys.Down ? 1 : -1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                RunSelected();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void MoveSelection(int delta)
        {
            if (_list.Items.Count == 0) return;
            int idx = _list.SelectedIndices.Count > 0 ? _list.SelectedIndices[0] : -1;
            int next = idx + delta;
            if (next < 0) next = 0;
            if (next > _list.Items.Count - 1) next = _list.Items.Count - 1;
            _list.Items[next].Selected = true;
            _list.Items[next].Focused = true;
            _list.EnsureVisible(next);
        }

        private void RunSelected()
        {
            if (_list.SelectedItems.Count == 0) return;
            SelectedCommandId = _list.SelectedItems[0].Tag as string;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
