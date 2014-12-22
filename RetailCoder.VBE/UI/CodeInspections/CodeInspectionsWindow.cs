using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeInspections
{
    [ComVisible(false)]
    public partial class CodeInspectionsWindow : UserControl, IDockableUserControl
    {
        private const string ClassId = "D3B2A683-9856-4246-BDC8-6B0795DC875B";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return "Code Inspections"; } }

        public CodeInspectionsWindow()
        {
            InitializeComponent();
            RefreshButton.Click += RefreshButtonClicked;
            QuickFixButton.ButtonClick += QuickFixButton_Click;
            GoButton.Click += GoButton_Click;
            PreviousButton.Click += PreviousButton_Click;
            NextButton.Click += NextButton_Click;

            var items = new List<CodeInspectionResultGridViewItem>();
            CodeIssuesGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            CodeIssuesGridView.DataSource = new BindingList<CodeInspectionResultGridViewItem>(items);
            CodeIssuesGridView.AutoResizeColumns();
            CodeIssuesGridView.Columns["Issue"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            CodeIssuesGridView.SelectionChanged += CodeIssuesGridView_SelectionChanged;
            CodeIssuesGridView.CellDoubleClick += CodeIssuesGridView_CellDoubleClick;
        }

        private void QuickFixButton_Click(object sender, EventArgs e)
        {
            QuickFixItemClick(QuickFixButton.DropDownItems.Cast<ToolStripMenuItem>().First(item => item.Checked), EventArgs.Empty);
        }

        private void PreviousButton_Click(object sender, EventArgs e)
        {
            var previousIssueIndex = (CodeIssuesGridView.SelectedRows[0].Index == 0)
                ? CodeIssuesGridView.Rows.Count - 1
                : CodeIssuesGridView.SelectedRows[0].Index - 1;

            CodeIssuesGridView.Rows[previousIssueIndex].Selected = true;
            var item = CodeIssuesGridView.Rows[previousIssueIndex].DataBoundItem as CodeInspectionResultGridViewItem;
            OnNavigateCodeIssue(item);
        }

        public void FindNextIssue()
        {
            NextButton_Click(this, EventArgs.Empty);
        }

        private void NextButton_Click(object sender, EventArgs e)
        {
            if (CodeIssuesGridView.Rows.Count == 0)
            {
                return;
            }

            var nextIssueIndex = (CodeIssuesGridView.SelectedRows[0].Index == CodeIssuesGridView.Rows.Count - 1)
                ? 0
                : CodeIssuesGridView.SelectedRows[0].Index + 1;

            CodeIssuesGridView.Rows[nextIssueIndex].Selected = true;
            var item = CodeIssuesGridView.Rows[nextIssueIndex].DataBoundItem as CodeInspectionResultGridViewItem;
            OnNavigateCodeIssue(item);
        }

        private IDictionary<string, Action<VBE>> _availableQuickFixes;
        private void CodeIssuesGridView_SelectionChanged(object sender, EventArgs e)
        {
            var enableNavigation = (CodeIssuesGridView.SelectedRows.Count != 0);
            NextButton.Enabled = enableNavigation;
            PreviousButton.Enabled = enableNavigation;
            GoButton.Enabled = enableNavigation;

            if (CodeIssuesGridView.SelectedRows.Count > 0)
            {
                var issue = (CodeInspectionResultGridViewItem) CodeIssuesGridView.SelectedRows[0].DataBoundItem;
                _availableQuickFixes = issue.GetInspectionResultItem()
                    .GetQuickFixes();
                var descriptions = _availableQuickFixes.Keys.ToList();

                var quickFixMenu = QuickFixButton.DropDownItems;
                if (quickFixMenu.Count > 0)
                {
                    foreach (var quickFixButton in quickFixMenu.Cast<ToolStripMenuItem>())
                    {
                        quickFixButton.Click -= QuickFixItemClick;
                    }
                }

                quickFixMenu.Clear();
                foreach (var caption in descriptions)
                {
                    var item = (ToolStripMenuItem) quickFixMenu.Add(caption);
                    if (quickFixMenu.Count > 0)
                    {
                        item.CheckOnClick = false;
                        item.Checked = quickFixMenu.Count == 1;
                        item.Click += QuickFixItemClick;
                    }
                }
            }

            QuickFixButton.Enabled = QuickFixButton.HasDropDownItems;
        }

        public event EventHandler<QuickFixEventArgs> QuickFix;
        private void QuickFixItemClick(object sender, EventArgs e)
        {
            var quickFixButton = (ToolStripMenuItem)sender;
            if (QuickFix == null)
            {
                return;
            }

            var args = new QuickFixEventArgs(_availableQuickFixes[quickFixButton.Text]);
            QuickFix(this, args);
        }

        public void SetContent(IEnumerable<CodeInspectionResultGridViewItem> inspectionResults)
        {
            var results = inspectionResults.ToList();
            StatusLabel.Text = string.Format("{0} issue" + (results.Count > 1 ? "s" : string.Empty), results.Count);

            CodeIssuesGridView.DataSource = new BindingList<CodeInspectionResultGridViewItem>(results);
            CodeIssuesGridView.Refresh();
        }

        private void GoButton_Click(object sender, EventArgs e)
        {
            var issue = CodeIssuesGridView.SelectedRows[0].DataBoundItem as CodeInspectionResultGridViewItem;
            OnNavigateCodeIssue(issue);
        }

        public event EventHandler<NavigateCodeIssueEventArgs> NavigateCodeIssue;
        private void CodeIssuesGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }
            var issue = CodeIssuesGridView.Rows[e.RowIndex].DataBoundItem as CodeInspectionResultGridViewItem;
            OnNavigateCodeIssue(issue);
        }

        private void OnNavigateCodeIssue(CodeInspectionResultGridViewItem item)
        {
            var handler = NavigateCodeIssue;
            if (handler == null)
            {
                return;
            }

            handler(this, new NavigateCodeIssueEventArgs(item.GetInspectionResultItem().Node));
        }

        public event EventHandler RefreshCodeInspections;
        private void RefreshButtonClicked(object sender, EventArgs e)
        {
            var handler = RefreshCodeInspections;
            if (handler == null)
            {
                return;
            }

            handler(this, EventArgs.Empty);
        }
    }
}
