using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteIfBlock : AutoCompleteBlockBase
    {
        public AutoCompleteIfBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.Then}", $"{Tokens.End} {Tokens.If}") { }

        protected override bool MatchInputTokenAtEndOfLineOnly => true;
    }
}
