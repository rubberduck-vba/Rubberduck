using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldView :IDialogView
    {
        string NewPropertyName { get; set; }
        string VariableName { get; set; }
        Declaration TargetDeclaration { get; set; }
        EncapsulateFieldDialog.Accessibility PropertyAccessibility { get; set; }
        EncapsulateFieldDialog.SetterType PropertySetterType { get; set; }
        bool IsPropertySetterTypeChangeable { get; set; }
    }
}