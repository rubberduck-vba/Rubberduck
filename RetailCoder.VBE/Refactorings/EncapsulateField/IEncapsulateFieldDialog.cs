using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldDialog :IDialogView
    {
        Declaration TargetDeclaration { get; set; }

        string NewPropertyName { get; set; }
        bool CanImplementLetSetterType { get; set; }
        bool CanImplementSetSetterType { get; set; }
        bool LetSetterSelected { get; }
        bool SetSetterSelected { get; }
        bool MustImplementLetSetterType { get; }
        bool MustImplementSetSetterType { get; }

        string ParameterName { get; set; }
    }
}
