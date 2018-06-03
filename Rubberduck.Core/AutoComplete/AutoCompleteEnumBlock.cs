using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteEnumBlock : AutoCompleteBlockBase
    {
        public AutoCompleteEnumBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"{Tokens.Enum}", $"{Tokens.End} {Tokens.Enum}") { }
    }
}
