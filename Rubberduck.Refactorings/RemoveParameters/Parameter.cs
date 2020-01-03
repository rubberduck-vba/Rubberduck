using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Refactorings.RemoveParameters
{
    /// <summary>
    /// Contains data about a method parameter for the Remove Parameters refactoring.
    /// </summary>
    public class Parameter
    {
        public ParameterDeclaration Declaration { get; }
        public string Name { get; }
        public bool IsParamArray => Declaration.IsParamArray;

        public Parameter(ParameterDeclaration declaration, bool isRemoved = false)
        {
            Declaration = declaration;
            Name = declaration.Context.GetText().RemoveExtraSpacesLeavingIndentation();
        }
    }
}
