using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteEnumBlock : AutoCompleteBlockBase
    {
        public AutoCompleteEnumBlock()
            : base($"{Tokens.Enum} ", $"{Tokens.End} {Tokens.Enum}") { }
    }
}
