using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteDoBlock : AutoCompleteBlockBase
    {
        public AutoCompleteDoBlock(IIndenterSettings indenterSettings)
            : base(indenterSettings, $"{Tokens.Do}", Tokens.Loop) { }
    }
}
