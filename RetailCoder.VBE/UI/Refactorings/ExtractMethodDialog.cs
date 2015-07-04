using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ExtractMethod;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractMethodDialog : Form, IExtractMethodDialog
    {
        public ExtractMethodDialog()
        {
            _parameters = new BindingList<ExtractedParameter>();

            InitializeComponent();
            Localize();
            RegisterViewEvents();

            MethodAccessibilityCombo.DataSource = new[]
            {
                Accessibility.Private,
                Accessibility.Public,
                Accessibility.Friend
            }.ToList();
        }

        private void Localize()
        {
            Text = RubberduckUI.ExtractMethod_Caption;
            OkButton.Text = RubberduckUI.OK;
            CancelButton.Text = RubberduckUI.CancelButtonText;

            TitleLabel.Text = RubberduckUI.ExtractMethod_TitleText;
            InstructionsLabel.Text = RubberduckUI.ExtractMethod_InstructionsText;
            NameLabel.Text = RubberduckUI.NameLabelText;
            ReturnLabel.Text = RubberduckUI.ExtractMethod_ReturnLabel;
            SetReturnValueCheck.Text = RubberduckUI.ExtractMethod_SetBoxLabel;
            AccessibilityLabel.Text = RubberduckUI.ExtractMethod_AccessibilityLabel;
            ParametersLabel.Text = RubberduckUI.ExtractMethod_ParametersLabel;
            PreviewLabel.Text = RubberduckUI.ExtractMethod_PreviewLabel;
        }

        private void InitializeParameterGrid()
        {
            MethodParametersGrid.AutoGenerateColumns = false;
            MethodParametersGrid.Columns.Clear();
            MethodParametersGrid.DataSource = _parameters;
            MethodParametersGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
            MethodParametersGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            var paramNameColumn = new DataGridViewTextBoxColumn();
            paramNameColumn.Name = "Name";
            paramNameColumn.DataPropertyName = "Name";
            paramNameColumn.HeaderText = RubberduckUI.Name;
            paramNameColumn.ReadOnly = true;

            var paramTypeColumn = new DataGridViewTextBoxColumn();
            paramTypeColumn.Name = "TypeName";
            paramTypeColumn.DataPropertyName = "TypeName";
            paramTypeColumn.HeaderText = RubberduckUI.Type;
            paramTypeColumn.ReadOnly = true;

            var paramPassedByColumn = new DataGridViewTextBoxColumn();
            paramPassedByColumn.Name = "Passed";
            paramPassedByColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            paramPassedByColumn.HeaderText = RubberduckUI.Passed;
            paramPassedByColumn.DataPropertyName = "Passed";
            paramPassedByColumn.ReadOnly = true;

            MethodParametersGrid.Columns.AddRange(paramNameColumn, paramTypeColumn, paramPassedByColumn);
        }

        private void RegisterViewEvents()
        {
            SetReturnValueCheck.CheckedChanged += SetReturnValueCheck_CheckedChanged;
            MethodNameBox.TextChanged += MethodNameBox_TextChanged;
            MethodAccessibilityCombo.SelectedIndexChanged += MethodAccessibilityCombo_SelectedIndexChanged;
            MethodReturnValueCombo.SelectedIndexChanged += MethodReturnValueCombo_SelectedIndexChanged;
        }

        public event EventHandler RefreshPreview;
        public void OnRefreshPreview()
        {
            OnViewEvent(RefreshPreview);
        }

        private bool _setReturnValue;

        public bool SetReturnValue
        {
            get { return _setReturnValue; }
            set
            {
                _setReturnValue = value;
                OnRefreshPreview();
            }
        }

        public bool CanSetReturnValue
        {
            get { return SetReturnValueCheck.Enabled; }
            set
            {
                SetReturnValueCheck.Enabled = value;
                SetReturnValueCheck.Checked = value;
            }
        }

        private void SetReturnValueCheck_CheckedChanged(object sender, EventArgs e)
        {
            SetReturnValue = SetReturnValueCheck.Checked;
        }

        private void MethodReturnValueCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReturnValue = (ExtractedParameter) MethodReturnValueCombo.SelectedItem;
        }

        private Accessibility _accessibility;
        public Accessibility Accessibility
        {
            get { return _accessibility; }
            set
            {
                _accessibility = value; 
                OnRefreshPreview();
            }
        }

        private void MethodAccessibilityCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            Accessibility = ((Accessibility) MethodAccessibilityCombo.SelectedItem);
        }

        private void MethodNameBox_TextChanged(object sender, EventArgs e)
        {
            MethodName = MethodNameBox.Text;
            OnRefreshPreview();
        }

        private void OnViewEvent(EventHandler target, EventArgs args = null)
        {
            var handler = target;
            if (handler == null)
            {
                return;
            }

            handler(this, args ?? EventArgs.Empty);
        }

        public event EventHandler CancelButtonClicked;
        public void OnCancelButtonClicked(object sender, EventArgs e)
        {
            OnViewEvent(CancelButtonClicked);
        }

        public event EventHandler OkButtonClicked;
        public void OnOkButtonClicked(object sender, EventArgs e)
        {
            OnViewEvent(OkButtonClicked);
        }

        private string _preview;
        public string Preview
        {
            get { return _preview; }
            set
            {
                _preview = value;
                PreviewBox.Text = _preview;
            }
        }

        private BindingList<ExtractedParameter> _parameters;
        public IEnumerable<ExtractedParameter> Parameters
        {
            get { return _parameters.Where(p => p.Name != _returnValue.Name); }
            set
            {
                _parameters = new BindingList<ExtractedParameter>(value.ToList());
                InitializeParameterGrid();
                OnRefreshPreview();
            }
        }

        private BindingList<ExtractedParameter> _returnValues; 
        public IEnumerable<ExtractedParameter> ReturnValues
        {
            get { return _returnValues; }
            set
            {
                _returnValues = new BindingList<ExtractedParameter>(value.ToList());
                var items = _returnValues.ToArray();
                MethodReturnValueCombo.Items.AddRange(items);
                MethodReturnValueCombo.DisplayMember = "Name";
            }
        }

        private ExtractedParameter _returnValue;
        public ExtractedParameter ReturnValue
        {
            get { return _returnValue; }
            set
            {
                _returnValue = value;
                MethodReturnValueCombo.Text = value.Name;
                Parameters = Inputs.Where(input => input.Name != _returnValue.Name)
                                   .Union(Outputs.Where(output => output.Name != _returnValue.Name));
                OnRefreshPreview();
            }
        }

        public IEnumerable<ExtractedParameter> Inputs { get; set; }
        public IEnumerable<ExtractedParameter> Outputs { get; set; }
        public IEnumerable<ExtractedParameter> Locals { get; set; }

        public string MethodName
        {
            get { return MethodNameBox.Text; }
            set
            {
                MethodNameBox.Text = value;
                InvalidNameValidationIcon.Visible = string.IsNullOrWhiteSpace(value);
                OnRefreshPreview();
            }
        }
    }
}
