using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchReplaceModel
    {
        public VBE VBE { get; private set; }
        public VBProjectParseResult ParseResult { get; private set; }

        public RegexSearchReplaceModel(VBE vbe, VBProjectParseResult parseResult)
        {
            VBE = vbe;
            ParseResult = parseResult;
        }
    }
}