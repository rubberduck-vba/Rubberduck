using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteForBlock : AutoCompleteBlockBase
    {
        public AutoCompleteForBlock(IIndenterSettings indenterSettings) 
            : base(indenterSettings, $"{Tokens.For}", Tokens.Next) { }
    }
}
