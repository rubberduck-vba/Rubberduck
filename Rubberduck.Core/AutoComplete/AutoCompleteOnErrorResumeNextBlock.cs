using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteOnErrorResumeNextBlock : AutoCompleteBlockBase
    {
        public AutoCompleteOnErrorResumeNextBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings)
            : base(api, indenterSettings, $"{Tokens.On} {Tokens.Error} {Tokens.Resume} {Tokens.Next}", $"{Tokens.On} {Tokens.Error} {Tokens.GoTo} 0") { }

        protected override bool ExecuteOnCommittedInputOnly => false;
    }
}
