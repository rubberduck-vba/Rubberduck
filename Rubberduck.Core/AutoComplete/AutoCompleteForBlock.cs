using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteForBlock : AutoCompleteBlockBase
    {
        public AutoCompleteForBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings) 
            : base(api, indenterSettings, $"{Tokens.For}", Tokens.Next) { }
    }
}
