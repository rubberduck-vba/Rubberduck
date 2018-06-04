using Rubberduck.Parsing.Grammar;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.AutoComplete
{
    public class AutoCompleteEnumBlock : AutoCompleteBlockBase
    {
        public AutoCompleteEnumBlock(IVBETypeLibsAPI api, IIndenterSettings indenterSettings)
            : base(api, indenterSettings, $"{Tokens.Enum}", $"{Tokens.End} {Tokens.Enum}") { }
    }
}
