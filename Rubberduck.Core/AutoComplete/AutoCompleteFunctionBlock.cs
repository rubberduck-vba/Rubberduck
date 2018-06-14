using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteFunctionBlock : AutoCompleteBlockBase
    {
        public AutoCompleteFunctionBlock(IConfigProvider<IndenterSettings> indenterSettings) 
            : base(indenterSettings, $"{Tokens.Function}", $"{Tokens.End} {Tokens.Function}") { }
    }
}
