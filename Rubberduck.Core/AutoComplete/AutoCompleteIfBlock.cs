using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteIfBlock : AutoCompleteBlockBase
    {
        public AutoCompleteIfBlock() 
            : base($" {Tokens.Then}\n", $"{Tokens.End} {Tokens.If}") { }
    }
}
