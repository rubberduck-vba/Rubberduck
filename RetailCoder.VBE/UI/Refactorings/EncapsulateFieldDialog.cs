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
        
        public bool ImplementLetSetterType
        {
            get { return LetSetterTypeCheckBox.Checked; }
            set { LetSetterTypeCheckBox.Checked = value; }
        }

        public bool ImplementSetSetterType
        {
            get { return SetSetterTypeCheckBox.Checked; }
            set { SetSetterTypeCheckBox.Checked = value; }
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

            LetSetterTypeCheckBox.CheckedChanged += EncapsulateFieldDialog_SetterTypeChanged;
            SetSetterTypeCheckBox.CheckedChanged += EncapsulateFieldDialog_SetterTypeChanged;

            Shown += EncapsulateFieldDialog_Shown;
        }

        void EncapsulateFieldDialog_SetterTypeChanged(object sender, EventArgs e)
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
            ValidatePropertyName();
            ValidateVariableName();
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

            var getterText = string.Join(Environment.NewLine,
                string.Format("Public Property Get {0}() As {1}", NewPropertyName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", NewPropertyName, TargetDeclaration.IdentifierName),
                "End Property");

            var letterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + Environment.NewLine + "Public Property Let {0}(ByVal {1} As {2})",
                    NewPropertyName, ParameterName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", TargetDeclaration.IdentifierName, ParameterName),
                "End Property");

            var setterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + Environment.NewLine + "Public Property Set {0}(ByVal {1} As {2})",
                    NewPropertyName, ParameterName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", TargetDeclaration.IdentifierName, ParameterName),
                "End Property");

            PreviewBox.Text = getterText +
                              (ImplementLetSetterType ? letterText : string.Empty) +
                              (ImplementSetSetterType ? setterText : string.Empty);
        }

        private void ValidatePropertyName()
        {
            InvalidPropertyNameIcon.Visible = ValidateName(NewPropertyName, ParameterName);

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
                               || changedName == TargetDeclaration.IdentifierName
                               || changedName == otherName
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
