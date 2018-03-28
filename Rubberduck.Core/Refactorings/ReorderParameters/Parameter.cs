using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.ReorderParameters
{
    /// <summary>
    /// Contains data about a method parameter for the Reorder Parameters refactoring.
    /// </summary>
    public class Parameter
    {
        public string Name { get; }
        public Declaration Declaration { get; }
        public int Index { get;  }
        public bool IsOptional { get; }
        public bool IsParamArray { get; }

        public Parameter(ParameterDeclaration declaration, int index)
        {
            Declaration = declaration;
            Name = declaration.Context.GetText().RemoveExtraSpacesLeavingIndentation();
            Index = index;
            IsOptional = declaration.IsOptional;
            IsParamArray = declaration.IsParamArray;
        }
    }
}
