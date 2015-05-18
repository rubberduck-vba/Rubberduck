using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public partial class ReorderParametersDialog : Form, IReorderParametersView
    {
        public List<Parameter> Parameters { get; set; }
        private Parameter _selectedItem;

        public ReorderParametersDialog()
        {
            Parameters = new List<Parameter>();
            InitializeComponent();
            InitializeCaptions();

            MethodParametersGrid.SelectionChanged += MethodParametersGrid_SelectionChanged;
        }

        private void InitializeCaptions()
        {
            OkButton.Text = RubberduckUI.OkButtonText;
            CancelButton.Text = RubberduckUI.CancelButtonText;
            Text = RubberduckUI.ReorderParamsDialog_Caption;
            TitleLabel.Text = RubberduckUI.ReorderParamsDialog_TitleText;
            InstructionsLabel.Text = RubberduckUI.ReorderParamsDialog_InstructionsLabelText;
            MoveUpButton.Text = RubberduckUI.ReorderParamsDialog_MoveUpButtonText;
            MoveDownButton.Text = RubberduckUI.ReorderParamsDialog_MoveDownButtonText;
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

            var column = new DataGridViewTextBoxColumn
            {
                Name = "Parameter",
                DataPropertyName = "FullDeclaration",
                HeaderText = "Parameter",
                ReadOnly = true,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            };

            

            MethodParametersGrid.Columns.Add(column);
            _selectedItem = Parameters[0];
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

        private Declaration _target;
        public Declaration Target
        {
            get { return _target; }
            set { _target = value; }
        }

        private void MoveUpButtonClicked(object sender, EventArgs e)
        {
            if (MethodParametersGrid.SelectedRows.Count == 0)
            {
                return;
            }

            var selectedIndex = MethodParametersGrid.SelectedRows[0].Index;
            SwapParameters(selectedIndex, selectedIndex - 1);

            ReselectParameter();
        }

        private void MoveDownButtonClicked(object sender, EventArgs e)
        {
            if (MethodParametersGrid.SelectedRows.Count == 0)
            {
                return;
            }

            var selectedIndex = MethodParametersGrid.SelectedRows[0].Index;
            SwapParameters(selectedIndex, selectedIndex + 1);
            
            ReselectParameter();
        }

        private void SwapParameters(int index1, int index2)
        {
            var tmp = Parameters[index1];
            Parameters[index1] = Parameters[index2];
            Parameters[index2] = tmp;
        }

        private void ReselectParameter()
        {
            MethodParametersGrid.Refresh();
            MethodParametersGrid.Rows
                                .Cast<DataGridViewRow>()
                                .Single(row => row.DataBoundItem == _selectedItem).Selected = true;

            SelectionChanged();
        }

        private void SelectionChanged()
        {
            _selectedItem = MethodParametersGrid.SelectedRows.Count == 0
                ? null
                : (Parameter)MethodParametersGrid.SelectedRows[0].DataBoundItem;

            MoveUpButton.Enabled = _selectedItem != null
                && MethodParametersGrid.SelectedRows[0].Index != 0;

            MoveDownButton.Enabled = _selectedItem != null
                && MethodParametersGrid.SelectedRows[0].Index != Parameters.Count - 1;
        }
    }
}
