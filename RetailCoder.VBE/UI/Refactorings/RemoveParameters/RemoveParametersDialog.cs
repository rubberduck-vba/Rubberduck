using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactoring.RemoveParameterRefactoring;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public partial class RemoveParametersDialog : Form, IRemoveParametersView
    {
        public RemoveParameterRefactoring RemoveParams { get; set; }
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
            OkButton.Text = RubberduckUI.OkButtonText;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            Text = RubberduckUI.RemoveParamsDialog_Caption;
            TitleLabel.Text = RubberduckUI.RemoveParamsDialog_TitleText;
            InstructionsLabel.Text = RubberduckUI.RemoveParamsDialog_InstructionsLabelText;
            RemoveButton.Text = RubberduckUI.RemoveParamsDialog_RemoveButtonText;
            AddButton.Text = RubberduckUI.RemoveParamsDialog_AddButtonText;
        }

        private void MethodParametersGrid_SelectionChanged(object sender, EventArgs e)
        {
            SelectionChanged();
        }

        public void InitializeParameterGrid()
        {
            MethodParametersGrid.AutoGenerateColumns = false;
            MethodParametersGrid.Columns.Clear();
            MethodParametersGrid.DataSource = RemoveParams.Parameters;
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

        private void OkButtonClick(object sender, EventArgs e)
        {
            OnOkButtonClicked();
        }

        public event EventHandler CancelButtonClicked;
        public void OnCancelButtonClicked()
        {
            Hide();
        }

        public event EventHandler OkButtonClicked;
        public void OnOkButtonClicked()
        {
            var handler = OkButtonClicked;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void MarkToRemoveParam()
        {
            if (_selectedItem != null)
            {
                RemoveParams.Parameters.Find(item => item == _selectedItem).IsRemoved = true;
                SelectionChanged();
            }
        }

        private void MarkToAddParam()
        {
            if (_selectedItem != null)
            {
                RemoveParams.Parameters.Find(item => item == _selectedItem).IsRemoved = false;
                SelectionChanged();
            }
        }

        private void RemoveButtonClicked(object sender, EventArgs e)
        {
            MarkToRemoveParam();
        }

        private void AddButtonClicked(object sender, EventArgs e)
        {
            MarkToAddParam();
        }

        private void MethodParametersGrid_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (_selectedItem == null) { return; }
        
            if (_selectedItem.IsRemoved)
            {
                MarkToAddParam();
            }
            else
            {
                MarkToRemoveParam();
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
            AddButton.Enabled = _selectedItem != null && _selectedItem.IsRemoved;
        }
    }
}
