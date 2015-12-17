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
            get { return PropertyNameBox.Text; }
            set { PropertyNameBox.Text = value; }
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

            PropertyNameBox.TextChanged += PropertyNameBox_TextChanged;
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
            ValidateName();
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (TargetDeclaration == null) { return; }

            PreviewBox.Text = string.Join(Environment.NewLine,
                string.Format("Public Property Get {0}() As {1}", NewPropertyName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", NewPropertyName, TargetDeclaration.IdentifierName),
                "End Property" + Environment.NewLine,
                string.Format("Public Property {0} {1}({2} value As {3})", SetterTypeComboBox.SelectedValue,
                    NewPropertyName, VariableAccessibilityComboBox.SelectedValue, TargetDeclaration.AsTypeName),
                string.Format("    {0} = value", NewPropertyName),
                "End Property");
        }

        private void ValidateName()
        {
            if (TargetDeclaration == null) { return; }

            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            OkButton.Enabled = NewPropertyName != TargetDeclaration.IdentifierName
                               && char.IsLetter(NewPropertyName.FirstOrDefault())
                               && !tokenValues.Contains(NewPropertyName, StringComparer.InvariantCultureIgnoreCase)
                               && !NewPropertyName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameIcon.Visible = !OkButton.Enabled;
        }
    }
}
