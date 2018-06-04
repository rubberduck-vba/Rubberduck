using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWhileBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWhileBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"{Tokens.While}", Tokens.Wend) { }
    }
}
