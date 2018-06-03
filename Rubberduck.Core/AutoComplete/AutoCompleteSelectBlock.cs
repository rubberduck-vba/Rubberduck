using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteSelectBlock : AutoCompleteBlockBase
    {
        public AutoCompleteSelectBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"{Tokens.Select} {Tokens.Case}", $"{Tokens.End} {Tokens.Select}") { }
    }
}
