using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteTypeBlock : AutoCompleteBlockBase
    {
        public AutoCompleteTypeBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"{Tokens.Type}", $"{Tokens.End} {Tokens.Type}") { }
    }
}
