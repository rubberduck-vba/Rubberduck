using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteSubBlock : AutoCompleteBlockBase
    {
        public AutoCompleteSubBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.Sub}", $"{Tokens.End} {Tokens.Sub}") { }
    }
}
