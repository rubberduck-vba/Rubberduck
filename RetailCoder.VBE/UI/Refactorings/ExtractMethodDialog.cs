using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
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
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;

            TitleLabel.Text = RubberduckUI.ExtractMethod_TitleText;
            InstructionsLabel.Text = RubberduckUI.ExtractMethod_InstructionsText;
            NameLabel.Text = RubberduckUI.NameLabelText;
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

            var paramNameColumn = new DataGridViewTextBoxColumn
            {
                Name = "Name",
                DataPropertyName = "Name",
                HeaderText = RubberduckUI.Name,
                ReadOnly = true
            };

            var paramTypeColumn = new DataGridViewTextBoxColumn
            {
                Name = "TypeName",
                DataPropertyName = "TypeName",
                HeaderText = RubberduckUI.Type,
                ReadOnly = true
            };

            var paramPassedByColumn = new DataGridViewTextBoxColumn
            {
                Name = "Passed",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                HeaderText = RubberduckUI.Passed,
                DataPropertyName = "Passed",
                ReadOnly = true
            };

            MethodParametersGrid.Columns.AddRange(paramNameColumn, paramTypeColumn, paramPassedByColumn);
        }

        private void RegisterViewEvents()
        {
            MethodNameBox.TextChanged += MethodNameBox_TextChanged;
            MethodAccessibilityCombo.SelectedIndexChanged += MethodAccessibilityCombo_SelectedIndexChanged;
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
            ValidateName();
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
            }
        }

        private ExtractedParameter _returnValue;

        public IEnumerable<ExtractedParameter> Inputs { get; set; }
        public IEnumerable<ExtractedParameter> Outputs { get; set; }
        public IEnumerable<ExtractedParameter> Locals { get; set; }

        public string OldMethodName { get; set; }

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

        private void ValidateName()
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            OkButton.Enabled = MethodName != OldMethodName
                               && char.IsLetter(MethodName.FirstOrDefault())
                               && !tokenValues.Contains(MethodName, StringComparer.InvariantCultureIgnoreCase)
                               && !MethodName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameValidationIcon.Visible = !OkButton.Enabled;
        }

    }
}
