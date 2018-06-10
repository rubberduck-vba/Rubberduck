using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteEnumBlock : AutoCompleteBlockBase
    {
        public AutoCompleteEnumBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"{Tokens.Enum}", $"{Tokens.End} {Tokens.Enum}") { }

        protected override bool IndentBody => IndenterSettings.Create().IndentEnumTypeAsProcedure;
    }
}
