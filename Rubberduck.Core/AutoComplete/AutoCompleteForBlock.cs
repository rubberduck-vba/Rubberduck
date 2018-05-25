using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteForBlock : AutoCompleteBlockBase
    {
        public AutoCompleteForBlock() 
            : base($"{Tokens.For} ", Tokens.Next) { }
    }
}
