using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.UI.Refactorings
{
    public partial class EncapsulateFieldDialog : Form, IEncapsulateFieldView
    {
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
        
        public bool SetterTypeIsLet
        {
            get { return LetSetterTypeRadioButton.Checked; }
            set
            {
                if (value)
                {
                    LetSetterTypeRadioButton.Checked = true;
                }
                else
                {
                    SetSetterTypeRadioButton.Checked = true;
                }
            }
        }

        public bool IsSetterTypeChangeable
        {
            get { return SetterTypeGroupBox.Enabled; }
            set { SetterTypeGroupBox.Enabled = value; }
        }

        public EncapsulateFieldDialog()
        {
            InitializeComponent();
            LocalizeDialog();

            PropertyNameTextBox.TextChanged += PropertyNameBox_TextChanged;
            ParameterNameTextBox.TextChanged += VariableNameBox_TextChanged;
            ((RadioButton)SetterTypeGroupBox.Controls[0]).CheckedChanged += EncapsulateFieldDialog_CheckedChanged;

            Shown += EncapsulateFieldDialog_Shown;
        }

        void EncapsulateFieldDialog_CheckedChanged(object sender, EventArgs e)
        {
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
            UpdatePreview();
        }

        private void PropertyNameBox_TextChanged(object sender, EventArgs e)
        {
            ValidatePropertyName();
            UpdatePreview();
        }

        private void VariableNameBox_TextChanged(object sender, EventArgs e)
        {
            ValidateVariableName();
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (TargetDeclaration == null) { return; }

            PreviewBox.Text = string.Join(Environment.NewLine,
                string.Format("Public Property Get {0}() As {1}", NewPropertyName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", NewPropertyName, TargetDeclaration.IdentifierName),
                "End Property" + Environment.NewLine,
                string.Format("Public Property {0} {1}(ByVal {2} As {3})",
                    SetterTypeIsLet
                        ? LetSetterTypeRadioButton.Text
                        : SetSetterTypeRadioButton.Text,
                    NewPropertyName, ParameterName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", TargetDeclaration.IdentifierName, ParameterName),
                "End Property");
        }

        private void ValidatePropertyName()
        {
            if (TargetDeclaration == null) { return; }

            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            InvalidPropertyNameIcon.Visible = NewPropertyName == TargetDeclaration.IdentifierName
                               || !char.IsLetter(NewPropertyName.FirstOrDefault())
                               || tokenValues.Contains(NewPropertyName, StringComparer.InvariantCultureIgnoreCase)
                               || NewPropertyName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            SetOkButtonEnabledState();
        }

        private void ValidateVariableName()
        {
            if (TargetDeclaration == null) { return; }

            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            InvalidVariableNameIcon.Visible = ParameterName == TargetDeclaration.IdentifierName
                               || ParameterName == NewPropertyName
                               || !char.IsLetter(ParameterName.FirstOrDefault())
                               || tokenValues.Contains(ParameterName, StringComparer.InvariantCultureIgnoreCase)
                               || ParameterName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            SetOkButtonEnabledState();
        }

        private void SetOkButtonEnabledState()
        {
            OkButton.Enabled = !InvalidPropertyNameIcon.Visible && !InvalidVariableNameIcon.Visible;
        }
    }
}
