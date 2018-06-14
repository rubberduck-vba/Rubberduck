using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompletePropertyBlock : AutoCompleteBlockBase
    {
        public AutoCompletePropertyBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.Property}", $"{Tokens.End} {Tokens.Property}") { }
    }
}
