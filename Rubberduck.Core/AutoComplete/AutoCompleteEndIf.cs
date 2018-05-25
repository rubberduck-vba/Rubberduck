using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{

    public class AutoCompleteEndIf : AutoCompleteBlockBase
    {
        public AutoCompleteEndIf() 
            : base(Tokens.If, $"{Tokens.End} {Tokens.If}") { }
    }
}
