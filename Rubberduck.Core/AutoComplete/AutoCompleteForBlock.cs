using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteForBlock : AutoCompleteBlockBase
    {
        public AutoCompleteForBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.For}", Tokens.Next) { }
    }
}
