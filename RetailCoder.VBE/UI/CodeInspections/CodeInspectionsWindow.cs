using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Vbe.Interop;
using Rubberduck.Properties;

namespace Rubberduck.UI.CodeInspections
{
    public partial class CodeInspectionsWindow : UserControl, IDockableUserControl//, ICodeInspectionsWindow
    {
        private const string ClassId = "D3B2A683-9856-4246-BDC8-6B0795DC875B";
        string IDockableUserControl.ClassId { get { return ClassId; } }
        string IDockableUserControl.Caption { get { return RubberduckUI.CodeInspections; } }
        
        private BindingList<CodeInspectionResultGridViewItem> _inspectionResults;
        public BindingList<CodeInspectionResultGridViewItem> InspectionResults
        {
            get { return _inspectionResults; }
            set
            {
                _inspectionResults = value;
                CodeIssuesGridView.DataSource = _inspectionResults;
                CodeIssuesGridView.Refresh();
            }
        }

        public DataGridView GridView { get { return CodeIssuesGridView; } }

        public CodeInspectionsWindow()
        {
            InitializeComponent();
            InitWindow();
        }

        private void InitWindow()
        {
            RefreshButton.Click += RefreshButtonClicked;
            QuickFixButton.ButtonClick += QuickFixButton_Click;
            GoButton.Click += GoButton_Click;
            PreviousButton.Click += PreviousButton_Click;
            NextButton.Click += NextButton_Click;
            CopyButton.Click += CopyButton_Click;

            var items = new List<CodeInspectionResultGridViewItem>();
            CodeIssuesGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            CodeIssuesGridView.DataSource = new BindingList<CodeInspectionResultGridViewItem>(items);
            InspectionResults = CodeIssuesGridView.DataSource as BindingList<CodeInspectionResultGridViewItem>;

            CodeIssuesGridView.AutoResizeColumns();
            CodeIssuesGridView.Columns["Issue"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            CodeIssuesGridView.Columns["Severity"].Visible = false;

            CodeIssuesGridView.Columns["Icon"].HeaderText = RubberduckUI.Severity;
            CodeIssuesGridView.Columns["Project"].HeaderText = RubberduckUI.Project;
            CodeIssuesGridView.Columns["Component"].HeaderText = RubberduckUI.Component;
            CodeIssuesGridView.Columns["Line"].HeaderText = RubberduckUI.Line;
            CodeIssuesGridView.Columns["Issue"].HeaderText = RubberduckUI.Issue;

            CodeIssuesGridView.SelectionChanged += CodeIssuesGridView_SelectionChanged;
            CodeIssuesGridView.CellDoubleClick += CodeIssuesGridView_CellDoubleClick;

            QuickFixButton.Text = RubberduckUI.Fix;
            GoButton.Text = RubberduckUI.Go;
            NextButton.Text = RubberduckUI.Next;
            PreviousButton.Text = RubberduckUI.Previous;

            StatusLabel.Text = string.Format(RubberduckUI.CodeInspections_NumberOfIssues, 0, "s");

            RefreshButton.ToolTipText = RubberduckUI.CodeInspections_RefreshToolTip;
            QuickFixButton.ToolTipText = RubberduckUI.CodeInspections_QuickFixToolTip;
            GoButton.ToolTipText = RubberduckUI.CodeInspections_GoToolTip;
            PreviousButton.ToolTipText = RubberduckUI.CodeInspections_PreviousToolTip;
            NextButton.ToolTipText = RubberduckUI.CodeInspections_NextToolTip;
            CopyButton.ToolTipText = RubberduckUI.CodeInspections_CopyToolTip;
        }

        public void ToggleParsingStatus(bool enabled = true)
        {
            StatusLabel.Image = enabled
                ? Resources.hourglass
                : Resources.exclamation_diamond;
            StatusLabel.Text = enabled
                ? RubberduckUI.Parsing
                : RubberduckUI.CodeInspections_Inspecting;
        }

        public void SetIssuesStatus(int issueCount, bool completed = false)
        {
            _issueCount = issueCount;

            RefreshButton.Image = completed
                ? Resources.arrow_circle_double
                : Resources.cross_circle;

            if (!completed)
            {
                RefreshButton.Click -= RefreshButtonClicked;
                RefreshButton.Click += CancelButton_Click;
            }
            else
            {
                RefreshButton.Click -= CancelButton_Click;
                RefreshButton.Click += RefreshButtonClicked;
            }


            if (issueCount == 0)
            {
                if (completed)
                {
                    StatusLabel.Image = Resources.tick_circle;
                    StatusLabel.Text = RubberduckUI.OK;
                }
                else
                {
                    StatusLabel.Image = Resources.hourglass;
                    StatusLabel.Text = RubberduckUI.CodeInspections_Inspecting;
                }
            }
            else
            {
                if (completed)
                {
                    StatusLabel.Image = Resources.exclamation_diamond;
                    StatusLabel.Text = string.Format(RubberduckUI.CodeInspections_NumberOfIssues, issueCount, (issueCount != 1 ? "s" : string.Empty));
                }
                else
                {
                    StatusLabel.Image = Resources.hourglass;
                    StatusLabel.Text = string.Format(RubberduckUI.CodeInspections_InspectingIssues, RubberduckUI.CodeInspections_Inspecting, issueCount, (issueCount != 1 ? "s" : string.Empty));
                }
            }
        }

        public event EventHandler Cancel;
        private void CancelButton_Click(object sender, EventArgs e)
        {
            var handler = Cancel;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private int _issueCount;
        public void EnableRefresh(bool enabled = true)
        {
            RefreshButton.Enabled = enabled;
            QuickFixButton.Enabled = enabled && _issueCount > 0;
        }

        public event EventHandler CopyResults;
        private void CopyButton_Click(object sender, EventArgs e)
        {
            var handler = CopyResults;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
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

        private IDictionary<string, Action> _availableQuickFixes;
        private void CodeIssuesGridView_SelectionChanged(object sender, EventArgs e)
        {
            var enableNavigation = (CodeIssuesGridView.SelectedRows.Count != 0);
            NextButton.Enabled = enableNavigation;
            PreviousButton.Enabled = enableNavigation;
            GoButton.Enabled = enableNavigation;
            CopyButton.Enabled = enableNavigation;

            var quickFixMenu = QuickFixButton.DropDownItems;
            if (quickFixMenu.Count > 0)
            {
                foreach (var quickFixButton in quickFixMenu.Cast<ToolStripMenuItem>())
                {
                    quickFixButton.Click -= QuickFixItemClick;
                }
            }

            if (CodeIssuesGridView.SelectedRows.Count > 0)
            {
                var issue = (CodeInspectionResultGridViewItem) CodeIssuesGridView.SelectedRows[0].DataBoundItem;
                _availableQuickFixes = issue.GetInspectionResultItem()
                    .GetQuickFixes();
                var descriptions = _availableQuickFixes.Keys.ToList();

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

            InspectionResults = new BindingList<CodeInspectionResultGridViewItem>(results);
        }

        private void GoButton_Click(object sender, EventArgs e)
        {
            var issue = CodeIssuesGridView.SelectedRows[0].DataBoundItem as CodeInspectionResultGridViewItem;
            OnNavigateCodeIssue(issue);
        }

        public event EventHandler<NavigateCodeEventArgs> NavigateCodeIssue;
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

            var result = item.GetInspectionResultItem();
            handler(this, new NavigateCodeEventArgs(result.QualifiedSelection));
        }

        public event EventHandler RefreshCodeInspections;
        private void RefreshButtonClicked(object sender, EventArgs e)
        {
            var handler = RefreshCodeInspections;
            if (handler == null)
            {
                return;
            }

            toolStrip1.Refresh();

            handler(this, EventArgs.Empty);
        }

        public event EventHandler<DataGridViewCellMouseEventArgs> SortColumn;
        private void ColumnHeaderMouseClicked(object sender, DataGridViewCellMouseEventArgs e)
        {
            var handler = SortColumn;
            if (handler == null)
            {
                return;
            }

            handler(this, e);
        }
    }
}
