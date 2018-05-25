using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteTypeBlock : AutoCompleteBlockBase
    {
        public AutoCompleteTypeBlock()
            : base(Tokens.Type, $"{Tokens.End} {Tokens.Type}") { }
    }
}
