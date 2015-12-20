using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldView :IDialogView
    {
        Declaration TargetDeclaration { get; set; }

        string NewPropertyName { get; set; }
        bool IsSetterTypeChangeable { get; set; }

        bool ImplementLetSetterType { get; set; }
        bool ImplementSetSetterType { get; set; }

        string ParameterName { get; set; }
    }
}