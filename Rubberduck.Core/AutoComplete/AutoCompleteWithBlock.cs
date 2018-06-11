using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWithBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWithBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"{Tokens.With}", $"{Tokens.End} {Tokens.With}") { }
    }
}
