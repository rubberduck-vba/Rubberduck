using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Windows.Forms;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    [ExcludeFromCodeCoverage]
    public partial class ChangesControl : UserControl, IChangesView
    {
        public ChangesControl()
        {
            InitializeComponent();

            SetText();
        }

        private void SetText()
        {
            CurrentBranchLabel.Text = RubberduckUI.SourceControl_CurrentBranchLabel;
            CommitMessageLabel.Text = RubberduckUI.SourceControl_CommitMessageLabel;
            CommitButton.Text = RubberduckUI.SourceControl_CommitButtonLabel;
            IncludedChangesBox.Text = RubberduckUI.SourceControl_IncludedChanges;
            ExcludedChangesBox.Text = RubberduckUI.SourceControl_ExcludedChanges;
            UntrackedFilesBox.Text = RubberduckUI.SourceControl_UntrackedFiles;
        }

        public string CurrentBranch
        {
            get { return this.ChangesBranchNameLabel.Text; }
            set { this.ChangesBranchNameLabel.Text = value; }
        }

        public string CommitMessage
        {
            get { return this.CommitMessageBox.Text; }
            set { this.CommitMessageBox.Text = value; }
        }

        public CommitAction CommitAction
        {
            get { return (CommitAction)this.CommitActionDropdown.SelectedIndex; }
            set { this.CommitActionDropdown.SelectedIndex = (int)value; }
        }

        private BindingList<IFileStatusEntry> _includedChanges = new BindingList<IFileStatusEntry>();
        public IList<IFileStatusEntry> IncludedChanges
        {
            get { return _includedChanges; }
            set
            {
                _includedChanges = new BindingList<IFileStatusEntry>(value);
                this.IncludedChangesGrid.DataSource = _includedChanges;
                this.IncludedChangesGrid.Refresh();
            }
        }

        private BindingList<IFileStatusEntry> _excludedChanges = new BindingList<IFileStatusEntry>();
        public IList<IFileStatusEntry> ExcludedChanges
        {
            get { return _excludedChanges; }
            set
            {
                _excludedChanges = new BindingList<IFileStatusEntry>(value);
                this.ExcludedChangesGrid.DataSource = _excludedChanges;
                this.ExcludedChangesGrid.Refresh();
            }
        }

        private BindingList<IFileStatusEntry> _untrackedFiles = new BindingList<IFileStatusEntry>();
        public IList<IFileStatusEntry> UntrackedFiles
        {
            get { return _untrackedFiles; }
            set
            {
                _untrackedFiles = new BindingList<IFileStatusEntry>(value);
                this.UntrackedFilesGrid.DataSource = _untrackedFiles;
                this.UntrackedFilesGrid.Refresh();
            }
        }

        public bool CommitEnabled
        {
            get { return this.CommitButton.Enabled; }
            set { this.CommitButton.Enabled = value; }
        }

        public event EventHandler<EventArgs> SelectedActionChanged;
        private void CommitActionDropdown_SelectedIndexChanged(object sender, EventArgs e)
        {
            RaiseGenericEvent(SelectedActionChanged, e);
        }

        public event EventHandler<EventArgs> CommitMessageChanged;
        private void CommitMessageBox_TextChanged(object sender, EventArgs e)
        {
            RaiseGenericEvent(CommitMessageChanged, e);
        }

        public event EventHandler<EventArgs> Commit;
        private void CommitButton_Click(object sender, EventArgs e)
        {
            RaiseGenericEvent(Commit, e);
        }

        private void RaiseGenericEvent(EventHandler<EventArgs> handler, EventArgs e)
        {
            if (handler != null)
            {
                handler(this, e);
            }
        }

        #region DragDropHandling

        private Rectangle _dragBox;
        private int _row;
        private Control _dragSource;

        private void OnDataGridMouseDown(DataGridView sender, MouseEventArgs e)
        {
            _row = sender.HitTest(e.X, e.Y).RowIndex;
            if (_row != -1)
            {
                var dragSize = SystemInformation.DragSize;
                _dragBox = new Rectangle(
                            new Point(e.X - dragSize.Width / 2, e.Y - dragSize.Height / 2),
                            dragSize
                            );
            }
            else
            {
                _dragBox = Rectangle.Empty;
            }
        }

        private void OnDataGridMouseMove(DataGridView sender, MouseEventArgs e)
        {
            if (e.Button.HasFlag(MouseButtons.Left))
            {
                if (_dragBox != Rectangle.Empty && !_dragBox.Contains(e.X, e.Y))
                {
                    _dragSource = sender;
                    if (_dragSource == this.UntrackedFilesGrid)
                    {
                        this.ExcludedChangesGrid.AllowDrop = false;
                    }

                    try
                    {
                        DragDropEffects dropEffect = sender.DoDragDrop(sender.Rows[_row], DragDropEffects.Move);
                    }
                    finally
                    {
                        this.ExcludedChangesGrid.AllowDrop = true;
                        _dragSource = null;
                    }
                }
            }
        }

        private void ExcludedChangesGrid_DragDrop(object sender, DragEventArgs e)
        {
            if (_dragSource != null && _dragSource != this.UntrackedFilesGrid)
            {
                this.ExcludedChangesGrid.DataSource = MoveFileStatusItem(_excludedChanges, _includedChanges, _row);
            }

        }

        private void IncludedChangesGrid_DragDrop(object sender, DragEventArgs e)
        {
            if (_dragSource != null)
            {
                if (_dragSource == this.ExcludedChangesGrid)
                {
                    this.IncludedChangesGrid.DataSource = MoveFileStatusItem(_includedChanges, _excludedChanges, _row);
                }

                if (_dragSource == this.UntrackedFilesGrid)
                {
                    this.IncludedChangesGrid.DataSource = MoveFileStatusItem(_includedChanges, _untrackedFiles, _row);
                }
            }
        }

        private IList<IFileStatusEntry> MoveFileStatusItem(IList<IFileStatusEntry> destination,
            IList<IFileStatusEntry> source, int index)
        {
            destination.Add(source[index]);
            source.RemoveAt(index);

            return destination;
        }

        private void IncludedChangesGrid_MouseDown(object sender, MouseEventArgs e)
        {
            OnDataGridMouseDown((DataGridView)sender, e);
        }

        private void IncludedChangesGrid_MouseMove(object sender, MouseEventArgs e)
        {
            OnDataGridMouseMove((DataGridView)sender, e);
        }

        private void IncludedChangesGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ExcludedChangesGrid_MouseDown(object sender, MouseEventArgs e)
        {
            OnDataGridMouseDown((DataGridView)sender, e);
        }

        private void ExcludedChangesGrid_MouseMove(object sender, MouseEventArgs e)
        {
            OnDataGridMouseMove((DataGridView)sender, e);
        }

        private void ExcludedChangesGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void UntrackedFilesGrid_MouseDown(object sender, MouseEventArgs e)
        {
            OnDataGridMouseDown((DataGridView)sender, e);
        }

        private void UntrackedFilesGrid_MouseMove(object sender, MouseEventArgs e)
        {
            OnDataGridMouseMove((DataGridView)sender, e);
        }

        #endregion
    }
}
