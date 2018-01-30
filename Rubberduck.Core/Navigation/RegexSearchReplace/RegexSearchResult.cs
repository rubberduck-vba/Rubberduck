using System.Text.RegularExpressions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchResult
    {
        public Match Match { get; }
        public ICodeModule Module { get; }
        public Selection Selection { get; }
        public string DisplayString => Match.Value;

        public RegexSearchResult(Match match, ICodeModule module, int line)
        {
            Match = match;
            Module = module;
            Selection = new Selection(line, match.Index + 1, line, match.Index + match.Length + 1); // adjust columns for VBE 1-based indexing
        }
    }
}
