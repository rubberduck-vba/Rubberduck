using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteTypeBlock : AutoCompleteBlockBase
    {
        public AutoCompleteTypeBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings)
            : base(api, indenterSettings, $"{Tokens.Type}", $"{Tokens.End} {Tokens.Type}") { }
    }
}
