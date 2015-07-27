using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchResult
    {
        public Match Match { get; private set; }
        public CodeModule Module { get; private set; }
        public Selection Selection { get; private set; }

        public RegexSearchResult(Match match, CodeModule module, int line)
        {
            Match = match;
            Module = module;
            Selection = new Selection(line, match.Index, line, match.Index + match.Length);
        }
    }
}