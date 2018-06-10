using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteSelectBlock : AutoCompleteBlockBase
    {
        public AutoCompleteSelectBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"{Tokens.Select} {Tokens.Case}", $"{Tokens.End} {Tokens.Select}") { }
    }
}
