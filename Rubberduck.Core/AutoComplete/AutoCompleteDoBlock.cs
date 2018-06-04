using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteDoBlock : AutoCompleteBlockBase
    {
        public AutoCompleteDoBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings)
            : base(api, indenterSettings, $"{Tokens.Do}", Tokens.Loop) { }
    }
}
