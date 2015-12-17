using System;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing.Grammar;
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

        public string PropertyName
        {
            get { return PropertyNameBox.Text; }
            set
            {
                PropertyNameBox.Text = value;
                ValidateNewName();
            }
        }

        public Accessibility PropertyAccessibility
        {
            get { return (Accessibility)AccessibilityComboBox.SelectedItem; }
            set { AccessibilityComboBox.SelectedItem = value; }
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

            AccessibilityComboBox.DataSource = new[]
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

        private void PropertyNameBox_TextChanged(object sender, EventArgs e)
        {
            PropertyName = PropertyNameBox.Text;
        }

        private void ValidateNewName()
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);

            OkButton.Enabled = char.IsLetter(PropertyName.FirstOrDefault())
                               && !tokenValues.Contains(PropertyName, StringComparer.InvariantCultureIgnoreCase)
                               && !PropertyName.Any(c => !char.IsLetterOrDigit(c) && c != '_');

            InvalidNameIcon.Visible = !OkButton.Enabled;
        }
    }
}
