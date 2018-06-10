using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWhileBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWhileBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"{Tokens.While}", Tokens.Wend) { }
    }
}
