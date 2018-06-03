using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteIfBlock : AutoCompleteBlockBase
    {
        public AutoCompleteIfBlock(IIndenterSettings indenterSettings) 
            : base(indenterSettings, $"{Tokens.Then}", $"{Tokens.End} {Tokens.If}") { }
    }
}
