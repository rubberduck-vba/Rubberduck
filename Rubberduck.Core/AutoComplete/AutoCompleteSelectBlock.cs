using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteSelectBlock : AutoCompleteBlockBase
    {
        public AutoCompleteSelectBlock()
            : base($"{Tokens.Select} {Tokens.Case}", $"{Tokens.End} {Tokens.Select}") { }
    }
}
