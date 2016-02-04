using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldView :IDialogView
    {
        Declaration TargetDeclaration { get; set; }

        string NewPropertyName { get; set; }
        bool CanImplementLetSetterType { get; set; }
        bool CanImplementSetSetterType { get; set; }

        bool MustImplementLetSetterType { get; set; }
        bool MustImplementSetSetterType { get; set; }

        string ParameterName { get; set; }
    }
}