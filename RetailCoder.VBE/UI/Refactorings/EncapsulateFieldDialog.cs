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

            PropertyNameTextBox.TextChanged += PropertyNameBox_TextChanged;
            ParameterNameTextBox.TextChanged += VariableNameBox_TextChanged;
            
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
            if (MustImplementSetSetterType)
            {
                SetSetterTypeCheckBox.Checked = true;
                DisableAssignmentSelection();
            }
            else
            {
                LetSetterTypeCheckBox.Checked = true;
                if (MustImplementLetSetterType)
                {
                    DisableAssignmentSelection();
                }
                else
                {
                    LetSetterTypeCheckBox.CheckedChanged += EncapsulateFieldDialog_SetterTypeChanged;
                    SetSetterTypeCheckBox.CheckedChanged += EncapsulateFieldDialog_SetterTypeChanged;
                }
            }

            ValidatePropertyName();
            ValidateVariableName();
            UpdatePreview();
        }

        private void DisableAssignmentSelection()
        {
            LetSetterTypeCheckBox.Enabled = false;
            SetSetterTypeCheckBox.Enabled = false;            
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
            PreviewBox.Text = GetPropertyText();
        }

        private string GetPropertyText()
        {
            if (TargetDeclaration == null) { return string.Empty; }

            var getterText = GenerateGetter();

            var letterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + Environment.NewLine + "Public Property Let {0}(ByVal {1} As {2})",
                    NewPropertyName, ParameterName, TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", TargetDeclaration.IdentifierName, ParameterName),
                "End Property");

            var setterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + Environment.NewLine + "Public Property Set {0}(ByVal {1} As {2})",
                    NewPropertyName, ParameterName, TargetDeclaration.AsTypeName),
                string.Format("    Set {0} = {1}", TargetDeclaration.IdentifierName, ParameterName),
                "End Property");

            var propertyText = getterText +
                    (LetSetterTypeCheckBox.Checked ? letterText : string.Empty) +
                    (SetSetterTypeCheckBox.Checked ? setterText : string.Empty);

            var propertyTextLines = propertyText.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines));
        }

        private string GenerateGetter()
        {
            if (SetSetterTypeCheckBox.Checked && LetSetterTypeCheckBox.Checked)
            {
                return string.Join(Environment.NewLine,
                    string.Format("Public Property Get {0}() As {1}", NewPropertyName, TargetDeclaration.AsTypeName),
                    string.Format("    If VarType({0}) = vbObject Then", TargetDeclaration.IdentifierName),
                    string.Format("        Set {0} = {1}", NewPropertyName, TargetDeclaration.IdentifierName),
                                  "    Else",
                    string.Format("        {0} = {1}", NewPropertyName, TargetDeclaration.IdentifierName),
                                  "    End If",
                                  "End Property");
            }
            return string.Join(Environment.NewLine,
                string.Format("Public Property Get {0}() As {1}", NewPropertyName, TargetDeclaration.AsTypeName),
                string.Format("    {0}{1} = {2}", SetSetterTypeCheckBox.Checked ? "Set " : string.Empty, NewPropertyName, TargetDeclaration.IdentifierName),
                                "End Property");
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
