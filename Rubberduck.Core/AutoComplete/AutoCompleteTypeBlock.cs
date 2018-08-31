using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteTypeBlock : AutoCompleteBlockBase
    {
        public AutoCompleteTypeBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"{Tokens.Type}", $"{Tokens.End} {Tokens.Type}") { }
    }
}
