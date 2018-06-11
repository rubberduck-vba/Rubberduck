using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompletePrecompilerIfBlock : AutoCompleteBlockBase
    {
        public AutoCompletePrecompilerIfBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"#{Tokens.If}", $"#{Tokens.End} {Tokens.If}") { }

        protected override bool SkipPreCompilerDirective => false;
    }
}
