using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteDoBlock : AutoCompleteBlockBase
    {
        public AutoCompleteDoBlock()
            : base($"{Tokens.Do} ", Tokens.Loop) { }
    }
}
