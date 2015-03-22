using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    [SuppressMessage("ReSharper", "ArrangeThisQualifier")]
    public partial class SourceControlPanel : UserControl, ISourceControlView
    {
        public SourceControlPanel()
        {
            InitializeComponent();
        }

        public string ClassId
        {
            get { return "19A32FC9-4902-4385-9FE7-829D4F9C441D"; }
        }

        public string Caption
        {
            get { return "Source Control"; }
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

        //bug: control panel isn't repainting
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

        public event System.EventHandler<System.EventArgs> SelectedActionChanged;
        private void CommitActionDropdown_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            var handler = SelectedActionChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event System.EventHandler<System.EventArgs> CommitMessageChanged;
        private void CommitMessageBox_TextChanged(object sender, System.EventArgs e)
        {
            var handler = CommitMessageChanged;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event System.EventHandler<System.EventArgs> Commit;
        private void CommitButton_Click(object sender, System.EventArgs e)
        {
            var handler = Commit;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        public event System.EventHandler<System.EventArgs> RefreshData;
        private void RefreshButton_Click(object sender, System.EventArgs e)
        {
            var handler = RefreshData;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private Rectangle _dragBox;
        private int _row;

        private void IncludedChangesGrid_MouseDown(object sender, MouseEventArgs e)
        {
            OnDataGridMouseDown((DataGridView)sender, e);
        }

        private void OnDataGridMouseDown(DataGridView sender, MouseEventArgs e)
        {
            _row = sender.HitTest(e.X, e.Y).RowIndex;
            if (_row != -1)
            {
                var dragSize = SystemInformation.DragSize;
                _dragBox = new Rectangle(
                            new Point(e.X - dragSize.Width/2, e.Y - dragSize.Height/2),
                            dragSize
                            );
            }
            else
            {
                _dragBox = Rectangle.Empty;
            }
        }

        private void IncludedChangesGrid_MouseMove(object sender, MouseEventArgs e)
        {
            OnDataGridMouseMove((DataGridView)sender, e);
        }

        private void OnDataGridMouseMove(DataGridView sender, MouseEventArgs e)
        {
            if (e.Button.HasFlag(MouseButtons.Left))
            {
                if (_dragBox != Rectangle.Empty && !_dragBox.Contains(e.X, e.Y))
                {
                    DragDropEffects dropEffect = sender.DoDragDrop(sender.Rows[_row], DragDropEffects.Move);
                }
            }
        }

        private void ExcludedChangesGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ExcludedChangesGrid_DragDrop(object sender, DragEventArgs e)
        {
            _excludedChanges.Add(_includedChanges[_row]);
            this.ExcludedChangesGrid.DataSource = _excludedChanges;

            _includedChanges.RemoveAt(_row);
        }

        private void ExcludedChangesGrid_MouseDown(object sender, MouseEventArgs e)
        {
            OnDataGridMouseDown((DataGridView)sender, e);
        }

        private void IncludedChangesGrid_DragDrop(object sender, DragEventArgs e)
        {
            _includedChanges.Add(_excludedChanges[_row]);
            this.IncludedChangesGrid.DataSource = _includedChanges;
            
            _excludedChanges.RemoveAt(_row);
        }

        private void IncludedChangesGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void ExcludedChangesGrid_MouseMove(object sender, MouseEventArgs e)
        {
            OnDataGridMouseMove((DataGridView)sender, e);
        }

    }
}
