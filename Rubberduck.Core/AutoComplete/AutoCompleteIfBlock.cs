using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteIfBlock : AutoCompleteBlockBase
    {
        public AutoCompleteIfBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings) 
            : base(api, indenterSettings, $"{Tokens.Then}", $"{Tokens.End} {Tokens.If}") { }

        protected override bool MatchInputTokenAtEndOfLineOnly => true;
    }
}
