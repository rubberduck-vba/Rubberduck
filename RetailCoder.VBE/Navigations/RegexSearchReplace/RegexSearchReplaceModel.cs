using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchReplaceModel
    {
        public VBE VBE { get; private set; }
        public Selection Selection { get; private set; }

        public RegexSearchReplaceModel(VBE vbe, Selection selection)
        {
            VBE = vbe;
            Selection = selection;
        }
    }
}