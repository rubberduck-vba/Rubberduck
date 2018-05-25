using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWithBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWithBlock()
            : base(Tokens.With, $"{Tokens.End} {Tokens.With}") { }
    }
}
