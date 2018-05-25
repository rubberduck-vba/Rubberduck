using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteIfBlock : AutoCompleteBlockBase
    {
        public AutoCompleteIfBlock() 
            : base(Tokens.If, $"{Tokens.End} {Tokens.If}") { }
    }
}
