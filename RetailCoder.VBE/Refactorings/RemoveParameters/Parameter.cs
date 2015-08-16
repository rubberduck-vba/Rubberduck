using Rubberduck.Parsing.Symbols;
using Rubberduck.VBA;

namespace Rubberduck.Refactorings.RemoveParameters
{
    /// <summary>
    /// Contains data about a method parameter for the Remove Parameters refactoring.
    /// </summary>
    public class Parameter
    {
        public Declaration Declaration { get; private set; }
        public string Name { get; private set; }
        public int Index { get; private set; }
        public bool IsRemoved { get; set; }

        public Parameter(Declaration declaration, int index, bool isRemoved = false)
        {
            Declaration = declaration;
            Name = declaration.Context.GetText().RemoveExtraSpaces();
            Index = index;
            IsRemoved = isRemoved;
        }
    }
}
