using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplaceModel
    {
        public VBE VBE { get; private set; }
        public VBProjectParseResult ParseResult { get; private set; }
        public QualifiedSelection Selection { get; private set; }

        public RegexSearchReplaceModel(VBE vbe, VBProjectParseResult parseResult, QualifiedSelection selection)
        {
            VBE = vbe;
            ParseResult = parseResult;
            Selection = selection;
        }
    }
}