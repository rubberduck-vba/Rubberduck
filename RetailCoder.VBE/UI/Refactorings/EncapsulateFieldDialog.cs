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
        public enum Accessibility
        {
            ByVal,
            ByRef
        }

        public enum SetterType
        {
            Let,
            Set
        }
        
        public string NewPropertyName
        {
            get { return PropertyNameTextBox.Text; }
            set { PropertyNameTextBox.Text = value; }
        }

        public string VariableName
        {
            get { return VariableNameTextBox.Text; }
            set { VariableNameTextBox.Text = value; }
        }

        public Declaration TargetDeclaration { get; set; }

        public Accessibility PropertyAccessibility
        {
            get { return (Accessibility)VariableAccessibilityComboBox.SelectedItem; }
            set { VariableAccessibilityComboBox.SelectedItem = value; }
        }

        public SetterType PropertySetterType
        {
            get { return (SetterType)SetterTypeComboBox.SelectedItem; }
            set { SetterTypeComboBox.SelectedItem = value; }
        }

        public bool IsPropertySetterTypeChangeable
        {
            get { return SetterTypeComboBox.Enabled; }
            set { SetterTypeComboBox.Enabled = value; }
        }

        public EncapsulateFieldDialog()
        {
            InitializeComponent();
            LocalizeDialog();

            PropertyNameTextBox.TextChanged += PropertyNameBox_TextChanged;
            VariableNameTextBox.TextChanged += VariableNameBox_TextChanged;
            VariableAccessibilityComboBox.SelectedValueChanged += VariableAccessibilityComboBoxSelectedValueChanged;
            SetterTypeComboBox.SelectedValueChanged += SetterTypeComboBox_SelectedValueChanged;

            Shown += EncapsulateFieldDialog_Shown;

            VariableAccessibilityComboBox.DataSource = new[]
            {
                Accessibility.ByVal,
                Accessibility.ByRef,
            }.ToList();

            SetterTypeComboBox.DataSource = new[]
            {
                SetterType.Let,
                SetterType.Set,
            }.ToList();
        }

        private void LocalizeDialog()
        {
            Text = RubberduckUI.EncapsulateField_Caption;
            TitleLabel.Text = RubberduckUI.EncapsulateField_TitleText;
            InstructionsLabel.Text = RubberduckUI.EncapsulateField_InstructionText;
            PropertyNameLabel.Text = RubberduckUI.EncapsulateField_PropertyName;
            SetterTypeLabel.Text = RubberduckUI.EncapsulateField_SetterType;
            VariableNameLabel.Text = RubberduckUI.EncapsulateField_VariableName;
            AccessibilityLabel.Text = RubberduckUI.EncapsulateField_VariableAccessibility;
            OkButton.Text = RubberduckUI.OK;
            CancelDialogButton.Text = RubberduckUI.CancelButtonText;
        }

        void EncapsulateFieldDialog_Shown(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        void SetterTypeComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        void VariableAccessibilityComboBoxSelectedValueChanged(object sender, EventArgs e)
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
                string.Format("Public Property {0} {1}({2} {3} As {4})", SetterTypeComboBox.SelectedValue,
                    NewPropertyName, VariableAccessibilityComboBox.SelectedValue, VariableName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = value", TargetDeclaration.IdentifierName),
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

            InvalidVariableNameIcon.Visible = VariableName == TargetDeclaration.IdentifierName
                               || VariableName == NewPropertyName
                               || !char.IsLetter(VariableName.FirstOrDefault())
                               || tokenValues.Contains(VariableName, StringComparer.InvariantCultureIgnoreCase)
                               || VariableName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            SetOkButtonEnabledState();
        }

        private void SetOkButtonEnabledState()
        {
            OkButton.Enabled = !InvalidPropertyNameIcon.Visible && !InvalidVariableNameIcon.Visible;
        }
    }
}
