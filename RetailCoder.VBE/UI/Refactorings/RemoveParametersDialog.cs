using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.UI.Refactorings
{
    public partial class RemoveParametersDialog : Form, IRemoveParametersView
    {
        public List<Parameter> Parameters { get; set; }
        private Parameter _selectedItem;

        public RemoveParametersDialog()
        {
            InitializeComponent();
            InitializeCaptions();

            MethodParametersGrid.SelectionChanged += MethodParametersGrid_SelectionChanged;
            MethodParametersGrid.CellMouseDoubleClick += MethodParametersGrid_CellMouseDoubleClick;
        }

        private void InitializeCaptions()
        {
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
            Text = RubberduckUI.RemoveParamsDialog_Caption;
            TitleLabel.Text = RubberduckUI.RemoveParamsDialog_TitleText;
            InstructionsLabel.Text = RubberduckUI.RemoveParamsDialog_InstructionsLabelText;
            RemoveButton.Text = RubberduckUI.Remove;
            RestoreButton.Text = RubberduckUI.Restore;
        }

        private void MethodParametersGrid_SelectionChanged(object sender, EventArgs e)
        {
            SelectionChanged();
        }

        public void InitializeParameterGrid()
        {
            MethodParametersGrid.AutoGenerateColumns = false;
            MethodParametersGrid.Columns.Clear();
            MethodParametersGrid.DataSource = Parameters;
            MethodParametersGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            MethodParametersGrid.MultiSelect = false;
            MethodParametersGrid.AllowUserToResizeRows = false;
            MethodParametersGrid.AllowDrop = true;
            MethodParametersGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;

            var column = new DataGridViewTextBoxColumn
            {
                Name = "Parameter",
                DataPropertyName = "Name",
                HeaderText = RubberduckUI.Parameter,
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };

            MethodParametersGrid.Columns.Add(column);
        }

        private void MarkAsRemovedParam()
        {
            if (_selectedItem != null)
            {
                var indexOfRemoved = Parameters.FindIndex(item => item == _selectedItem);

                Parameters.ElementAt(indexOfRemoved).IsRemoved = true;
                MethodParametersGrid.Rows[indexOfRemoved].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Strikeout);

                SelectionChanged();
            }
        }

        private void MarkAsRestoredParam()  // really just un-mark as removed, but [tag:naming-is-hard]
        {
            if (_selectedItem != null)
            {
                var indexOfRemoved = Parameters.FindIndex(item => item == _selectedItem);

                Parameters.ElementAt(indexOfRemoved).IsRemoved = false;
                MethodParametersGrid.Rows[indexOfRemoved].DefaultCellStyle.Font = new Font(this.Font, FontStyle.Regular);

                SelectionChanged();
            }
        }

        private void RemoveButtonClicked(object sender, EventArgs e)
        {
            MarkAsRemovedParam();
        }

        private void RestoreButtonClicked(object sender, EventArgs e)
        {
            MarkAsRestoredParam();
        }

        private void MethodParametersGrid_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_selectedItem == null) { return; }
        
            if (_selectedItem.IsRemoved)
            {
                MarkAsRestoredParam();
            }
            else
            {
                MarkAsRemovedParam();
            }
        }

        private int GetFirstSelectedRowIndex(int index)
        {
            return MethodParametersGrid.SelectedRows[index].Index;
        }

        private void SelectionChanged()
        {
            _selectedItem = MethodParametersGrid.SelectedRows.Count == 0
                ? null
                : (Parameter)MethodParametersGrid.SelectedRows[0].DataBoundItem;

            RemoveButton.Enabled = _selectedItem != null && !_selectedItem.IsRemoved;
            RestoreButton.Enabled = _selectedItem != null && _selectedItem.IsRemoved;
        }
    }
}
