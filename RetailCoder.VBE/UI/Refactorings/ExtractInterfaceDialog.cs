using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractInterfaceDialog : Form, IExtractInterfaceView
    {
        public string InterfaceName
        {
            get { return InterfaceNameBox.Text; }
            set { InterfaceNameBox.Text = value; }
        }

        private List<InterfaceMember> _members;
        public List<InterfaceMember> Members
        {
            get { return _members; }
            set
            {
                _members = value;
                InitializeParameterGrid();
            }
        }

        public ExtractInterfaceDialog()
        {
            InitializeComponent();

            InterfaceNameBox.TextChanged += InterfaceNameBox_TextChanged;
            SelectAllButton.Click += SelectAllButton_Click;
            DeselectAllButton.Click += DeselectAllButton_Click;
        }

        void InterfaceNameBox_TextChanged(object sender, EventArgs e)
        {
            ValidateNewName();
        }

        void SelectAllButton_Click(object sender, EventArgs e)
        {
            ToggleSelection(true);
        }

        void DeselectAllButton_Click(object sender, EventArgs e)
        {
            ToggleSelection(false);
        }

        private void InitializeParameterGrid()
        {
            InterfaceMembersGridView.AutoGenerateColumns = false;
            InterfaceMembersGridView.Columns.Clear();
            InterfaceMembersGridView.DataSource = Members;
            InterfaceMembersGridView.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            InterfaceMembersGridView.MultiSelect = false;

            var isSelected = new DataGridViewCheckBoxColumn
            {
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells,
                Name = "IsSelected",
                DataPropertyName = "IsSelected",
                HeaderText = string.Empty,
                ReadOnly = false
            };

            var signature = new DataGridViewTextBoxColumn
            {
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                Name = "Signature",
                DataPropertyName = "Signature",
                HeaderText = "Signature",
                ReadOnly = true
            };

            InterfaceMembersGridView.Columns.AddRange(isSelected, signature);
        }

        void ToggleSelection(bool state)
        {
            foreach (var row in InterfaceMembersGridView.Rows.Cast<DataGridViewRow>())
            {
                row.Cells["IsSelected"].Value = state;
            }
        }

        private void ValidateNewName()
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            OkButton.Enabled = char.IsLetter(InterfaceName.FirstOrDefault())
                               && !tokenValues.Contains(InterfaceName, StringComparer.InvariantCultureIgnoreCase)
                               && !InterfaceName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameValidationIcon.Visible = !OkButton.Enabled;
        }
    }
}
