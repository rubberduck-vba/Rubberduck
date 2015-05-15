using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public partial class ReorderParametersDialog : Form, IReorderParametersView
    {
        public List<Parameter> Parameters { get; set; }
        public Parameter SelectedItem { get; set; }

        public ReorderParametersDialog()
        {
            Parameters = new List<Parameter>();
            InitializeComponent();

            MethodParametersGrid.SelectionChanged += MethodParametersGrid_SelectionChanged;
        }

        private void MethodParametersGrid_SelectionChanged(object sender, EventArgs e)
        {
            DataGridView sentVal = sender as DataGridView;
            SelectedItem = Parameters.Where(item => item.Name == sentVal.CurrentCell.Value).First();
        }

        public void InitializeParameterGrid()
        {
            MethodParametersGrid.AutoGenerateColumns = false;
            MethodParametersGrid.Columns.Clear();
            MethodParametersGrid.DataSource = Parameters;
            MethodParametersGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            MethodParametersGrid.MultiSelect = false;

            var paramNameColumn = new DataGridViewTextBoxColumn();
            paramNameColumn.Name = "Name";
            paramNameColumn.DataPropertyName = "Name";
            paramNameColumn.HeaderText = "Name";
            paramNameColumn.ReadOnly = true;
            paramNameColumn.Width = 262;    // fits nice

            MethodParametersGrid.Columns.Add(paramNameColumn);
            SelectedItem = Parameters[0];
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
            set
            {
                _target = value;
                if (_target == null)
                {
                    return;
                }
            }
        }

        private void MoveUpButtonClicked(object sender, EventArgs e)
        {
            int selectedIndex = Parameters.IndexOf(SelectedItem);

            if (selectedIndex == 0)
            {
                return;
            }

            Parameter tmp = Parameters[selectedIndex];
            Parameters[selectedIndex] = Parameters[selectedIndex - 1];
            Parameters[selectedIndex - 1] = tmp;
            MethodParametersGrid.Refresh();
        }

        private void MoveDownButtonClicked(object sender, EventArgs e)
        {
            int selectedIndex = Parameters.IndexOf(SelectedItem);

            if (selectedIndex == Parameters.Count() - 1)
            {
                return;
            }

            Parameter tmp = Parameters[selectedIndex];
            Parameters[selectedIndex] = Parameters[selectedIndex + 1];
            Parameters[selectedIndex + 1] = tmp;
            MethodParametersGrid.Refresh();
        }

        private void RegisterViewEvents()
        {
            OkButton.Click += OkButtonClicked;
            CancelButton.Click += CancelButtonClicked;
            MoveUpButton.Click += MoveUpButtonClicked;
            MoveDownButton.Click += MoveDownButtonClicked;
        }
    }
}
