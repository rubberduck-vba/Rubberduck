using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractInterfaceDialog : Form, IExtractInterfaceDialog
    {
        public string InterfaceName
        {
            get { return InterfaceNameBox.Text; }
            set { InterfaceNameBox.Text = value; }
        }

        private IEnumerable<InterfaceMember> _members;
        public IEnumerable<InterfaceMember> Members
        {
            get { return _members; }
            set
            {
                _members = value;
                InitializeParameterGrid();
            }
        }

        public List<string> ComponentNames { get; set; }

        public ExtractInterfaceDialog()
        {
            InitializeComponent();
            Localize();

            InterfaceNameBox.TextChanged += InterfaceNameBox_TextChanged;
            InterfaceMembersGridView.CellValueChanged += InterfaceMembersGridView_CellValueChanged;
            SelectAllButton.Click += SelectAllButton_Click;
            DeselectAllButton.Click += DeselectAllButton_Click;
        }

        private void Localize()
        {
            Text = RubberduckUI.ExtractInterface_Caption;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;

            NameLabel.Text = RubberduckUI.NameLabelText;
            TitleLabel.Text = RubberduckUI.ExtractInterface_TitleLabel;
            InstructionsLabel.Text = RubberduckUI.ExtractInterface_InstructionLabel;
            DeselectAllButton.Text = RubberduckUI.DeselectAll_Button;
            SelectAllButton.Text = RubberduckUI.SelectAll_Button;
            MembersGroupBox.Text = RubberduckUI.ExtractInterface_MembersGroupBox;
        }

        private void InterfaceNameBox_TextChanged(object sender, EventArgs e)
        {
            ValidateNewName();
        }

        private void InterfaceMembersGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            _members.ElementAt(e.RowIndex).IsSelected =
                (bool) InterfaceMembersGridView.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
        }

        private void SelectAllButton_Click(object sender, EventArgs e)
        {
            ToggleSelection(true);
        }

        private void DeselectAllButton_Click(object sender, EventArgs e)
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
                Name = "Members",
                DataPropertyName = "FullMemberSignature",
                ReadOnly = true
            };

            InterfaceMembersGridView.Columns.AddRange(isSelected, signature);
        }

        private void ToggleSelection(bool state)
        {
            foreach (var row in InterfaceMembersGridView.Rows.Cast<DataGridViewRow>())
            {
                row.Cells["IsSelected"].Value = state;
            }
        }

        private void ValidateNewName()
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            OkButton.Enabled = !ComponentNames.Contains(InterfaceName)
                               && InterfaceName.Length > 1
                               && char.IsLetter(InterfaceName.FirstOrDefault())
                               && !tokenValues.Contains(InterfaceName, StringComparer.InvariantCultureIgnoreCase)
                               && !InterfaceName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameValidationIcon.Visible = !OkButton.Enabled;
        }
    }
}
