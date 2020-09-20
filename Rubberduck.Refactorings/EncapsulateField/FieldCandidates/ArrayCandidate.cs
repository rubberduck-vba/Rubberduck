using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class ArrayCandidate : EncapsulateFieldCandidate
    {
        public ArrayCandidate(Declaration declaration)
            :base(declaration)
        {
            PropertyAsTypeName = Tokens.Variant;
            IsReadOnly = true;
        }
    }
}
