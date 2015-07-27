using System.Text.RegularExpressions;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchResult
    {
        public Match Match { get; private set; }
        public Selection Selection { get; private set; }

        public RegexSearchResult(Match match, int line)
        {
            Match = match;
            Selection = new Selection(line, match.Index, line, match.Index + match.Length);
        }
    }
}