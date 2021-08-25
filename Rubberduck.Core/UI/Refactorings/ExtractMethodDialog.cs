using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.Resources;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.UI.Refactorings
{
    public partial class ExtractMethodDialog : Form, IExtractMethodDialog
    {
        public ExtractMethodDialog()
        {
            _returnValue = null;
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
            Text = RefactoringsUI.ExtractMethod_Caption;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;

            TitleLabel.Text = RefactoringsUI.ExtractMethod_TitleText;
            InstructionsLabel.Text = RefactoringsUI.ExtractMethod_InstructionsText;
            NameLabel.Text = RubberduckUI.NameLabelText;
            AccessibilityLabel.Text = RefactoringsUI.ExtractMethod_AccessibilityLabel;
            ParametersLabel.Text = RefactoringsUI.ExtractMethod_ParametersLabel;
            PreviewLabel.Text = RefactoringsUI.ExtractMethod_PreviewLabel;
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
            get => _setReturnValue;
            set
            {
                _setReturnValue = value;
                OnRefreshPreview();
            }
        }



        private Accessibility _accessibility;
        public Accessibility Accessibility
        {
            get => _accessibility;
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

            handler?.Invoke(this, args ?? EventArgs.Empty);
        }

        private string _preview;
        public string Preview
        {
            get => _preview;
            set
            {
                _preview = value;
                PreviewBox.Text = _preview;
            }
        }

        private BindingList<ExtractedParameter> _parameters;
        public IEnumerable<ExtractedParameter> Parameters
        {
            get { return _parameters.Where(p => p.Name != _returnValue?.Name); }
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
            get => _returnValues;
            set => _returnValues = new BindingList<ExtractedParameter>(value.ToList());
        }

        private readonly ExtractedParameter _returnValue;

        public IEnumerable<ExtractedParameter> Inputs { get; set; }
        public IEnumerable<ExtractedParameter> Outputs { get; set; }
        public IEnumerable<ExtractedParameter> Locals { get; set; }

        public string OldMethodName { get; set; }

        public string MethodName
        {
            get => MethodNameBox.Text;
            set
            {
                MethodNameBox.Text = value;
                InvalidNameValidationIcon.Visible = string.IsNullOrWhiteSpace(value);
                OnRefreshPreview();
            }
        }

        private void ValidateName()
        {
            OkButton.Enabled = MethodName != OldMethodName
                               && char.IsLetter(MethodName.FirstOrDefault())
                               && !Tokens.IllegalIdentifierNames.Contains(MethodName, StringComparer.InvariantCultureIgnoreCase)
                               && !MethodName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameValidationIcon.Visible = !OkButton.Enabled;
        }

    }
}
