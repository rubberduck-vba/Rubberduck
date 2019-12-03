using System.Text.RegularExpressions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchResult
    {
        public Match Match { get; }

        //FIXME: We should not save COM wrappers anywhere.
        public ICodeModule Module { get; }
        public Selection Selection { get; }
        public string DisplayString => Match.Value;

        public RegexSearchResult(Match match, ICodeModule module, int line, int columnOffset = 0)
        {
            Match = match;
            Module = module;
            Selection = new Selection(line, match.Index + columnOffset + 1, line, match.Index + match.Length + columnOffset  + 1); // adjust columns for VBE 1-based indexing
        }
    }
}
