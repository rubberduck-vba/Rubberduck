using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteSelectBlock : AutoCompleteBlockBase
    {
        public AutoCompleteSelectBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings)
            : base(api, indenterSettings, $"{Tokens.Select} {Tokens.Case}", $"{Tokens.End} {Tokens.Select}") { }
    }
}
