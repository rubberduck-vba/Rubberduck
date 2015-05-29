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

        public RemoveParametersDialog()
        {
            InitializeComponent();
            InitializeCaptions();

            MethodParametersGrid.SelectionChanged += MethodParametersGrid_SelectionChanged;
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

        private void RemoveButtonClicked(object sender, EventArgs e)
        {
            
        }

        private void AddButtonClicked(object sender, EventArgs e)
        {
            
        }

        private int GetFirstSelectedRowIndex(int index)
        {
            return MethodParametersGrid.SelectedRows[index].Index;
        }

        /*private void ReselectParameter()
        {
            MethodParametersGrid.Refresh();
            MethodParametersGrid.Rows
                                .Cast<DataGridViewRow>()
                                .Single(row => row.DataBoundItem == _selectedItem).Selected = true;

            SelectionChanged();
        }*/

        private void SelectionChanged()
        {
           /*MoveUpButton.Enabled = _selectedItem != null
                && MethodParametersGrid.SelectedRows[0].Index != 0;

            MoveDownButton.Enabled = _selectedItem != null
                && MethodParametersGrid.SelectedRows[0].Index != RemoveParams.Parameters.Count - 1;*/
        }
    }
}
