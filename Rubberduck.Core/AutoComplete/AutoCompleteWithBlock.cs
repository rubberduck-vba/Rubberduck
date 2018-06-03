using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWithBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWithBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"{Tokens.With}", $"{Tokens.End} {Tokens.With}") { }
    }
}
