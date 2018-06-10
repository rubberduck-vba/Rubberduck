using Rubberduck.Parsing.Grammar;
using Rubberduck.SettingsProvider;
using Rubberduck.SmartIndenter;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteOnErrorResumeNextBlock : AutoCompleteBlockBase
    {
        public AutoCompleteOnErrorResumeNextBlock(IConfigProvider<IndenterSettings> indenterSettings)
            : base(indenterSettings, $"{Tokens.On} {Tokens.Error} {Tokens.Resume} {Tokens.Next}", $"{Tokens.On} {Tokens.Error} {Tokens.GoTo} 0") { }

        protected override bool ExecuteOnCommittedInputOnly => false;
    }
}
