using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteWhileBlock : AutoCompleteBlockBase
    {
        public AutoCompleteWhileBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings)
            : base(api, indenterSettings, $"{Tokens.While}", Tokens.Wend) { }
    }
}
