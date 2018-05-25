using Rubberduck.Parsing.Grammar;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWhileBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWhileBlock()
            : base($"{Tokens.While} ", Tokens.Wend) { }
    }
}
