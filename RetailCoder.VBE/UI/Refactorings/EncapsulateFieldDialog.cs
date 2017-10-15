using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings
{
    using SmartIndenter;

    public partial class EncapsulateFieldDialog : Form, IEncapsulateFieldDialog
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private PropertyGenerator _previewGenerator; 

        public string NewPropertyName
        {
            get { return PropertyNameTextBox.Text; }
            set { PropertyNameTextBox.Text = value; }
        }

        public string ParameterName
        {
            get { return ParameterNameTextBox.Text; }
            set { ParameterNameTextBox.Text = value; }
        }

        public Declaration TargetDeclaration { get; set; }

        public bool CanImplementLetSetterType { get; set; }

        public bool CanImplementSetSetterType { get; set; }

        public bool LetSetterSelected { get { return LetSetterTypeCheckBox.Checked; } }

        public bool SetSetterSelected { get { return SetSetterTypeCheckBox.Checked; } }

        public bool MustImplementLetSetterType
        {
            get { return CanImplementLetSetterType && !CanImplementSetSetterType; }
        }

        public bool MustImplementSetSetterType
        {
            get { return CanImplementSetSetterType && !CanImplementLetSetterType; }
        }

        public EncapsulateFieldDialog(RubberduckParserState state, IIndenter indenter)
        {
            _state = state;
            _indenter = indenter;

            InitializeComponent();
            LocalizeDialog();
            
            Shown += EncapsulateFieldDialog_Shown;
        }

        void EncapsulateFieldDialog_SetterTypeChanged(object sender, EventArgs e)
        {
            _previewGenerator.GenerateSetter = SetSetterTypeCheckBox.Checked;
            _previewGenerator.GenerateLetter = LetSetterTypeCheckBox.Checked;
            UpdatePreview();
        }

        private void LocalizeDialog()
        {
            Text = RubberduckUI.EncapsulateField_Caption;
            TitleLabel.Text = RubberduckUI.EncapsulateField_TitleText;
            InstructionsLabel.Text = RubberduckUI.EncapsulateField_InstructionText;
            PropertyNameLabel.Text = RubberduckUI.EncapsulateField_PropertyName;
            SetterTypeGroupBox.Text = RubberduckUI.EncapsulateField_SetterType;
            VariableNameLabel.Text = RubberduckUI.EncapsulateField_ParameterName;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
        }

        void EncapsulateFieldDialog_Shown(object sender, EventArgs e)
        {
            if (MustImplementSetSetterType)
            {
                SetSetterTypeCheckBox.Checked = true;
                LetSetterTypeCheckBox.Enabled = false;
            }
            else
            {
                LetSetterTypeCheckBox.Checked = true;
                SetSetterTypeCheckBox.Enabled = !MustImplementLetSetterType;
            }

            ValidatePropertyName();
            ValidateVariableName();

            _previewGenerator = new PropertyGenerator
            {
                PropertyName = NewPropertyName,
                AsTypeName = TargetDeclaration.AsTypeName,
                BackingField = TargetDeclaration.IdentifierName,
                ParameterName = ParameterName,
                GenerateSetter = SetSetterTypeCheckBox.Checked,
                GenerateLetter = LetSetterTypeCheckBox.Checked
            };

            LetSetterTypeCheckBox.CheckedChanged += EncapsulateFieldDialog_SetterTypeChanged;
            SetSetterTypeCheckBox.CheckedChanged += EncapsulateFieldDialog_SetterTypeChanged;
            PropertyNameTextBox.TextChanged += PropertyNameBox_TextChanged;
            ParameterNameTextBox.TextChanged += VariableNameBox_TextChanged;

            UpdatePreview();
        }

        private void PropertyNameBox_TextChanged(object sender, EventArgs e)
        {
            ValidatePropertyName();
            _previewGenerator.PropertyName = NewPropertyName;
            UpdatePreview();
        }

        private void VariableNameBox_TextChanged(object sender, EventArgs e)
        {
            ValidateVariableName();
            _previewGenerator.ParameterName = ParameterName;
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (TargetDeclaration == null)
            {
                PreviewBox.Text = string.Empty;
            }
            var propertyTextLines = _previewGenerator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            PreviewBox.Text = string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }

        private void ValidatePropertyName()
        {
            InvalidPropertyNameIcon.Visible = ValidateName(NewPropertyName, ParameterName) ||
                                              _state.AllUserDeclarations.Where(a => a.ParentScope == TargetDeclaration.ParentScope)
                                                                        .Any(a => a.IdentifierName.Equals(NewPropertyName, StringComparison.InvariantCultureIgnoreCase));

            SetOkButtonEnabledState();
        }

        private void ValidateVariableName()
        {
            InvalidVariableNameIcon.Visible = ValidateName(ParameterName, NewPropertyName);

            SetOkButtonEnabledState();
        }

        private bool ValidateName(string changedName, string otherName)
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            return TargetDeclaration == null
                               || changedName.Equals(TargetDeclaration.IdentifierName, StringComparison.InvariantCultureIgnoreCase)
                               || changedName.Equals(otherName, StringComparison.InvariantCultureIgnoreCase)
                               || !char.IsLetter(changedName.FirstOrDefault())
                               || tokenValues.Contains(ParameterName, StringComparer.InvariantCultureIgnoreCase)
                               || changedName.Any(c => !char.IsLetterOrDigit(c) && c != '_');
        }

        private void SetOkButtonEnabledState()
        {
            OkButton.Enabled = !InvalidPropertyNameIcon.Visible && !InvalidVariableNameIcon.Visible;
        }
    }
}
