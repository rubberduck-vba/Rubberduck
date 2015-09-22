using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public interface IRegexSearchReplace
    {
        IEnumerable<RegexSearchResult> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void Replace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void ReplaceAll(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
    }

    public class RegexSearchResult
    {
        public Match Match { get; private set; }
        public CodeModule Module { get; private set; }
        public Selection Selection { get; private set; }
        public string DisplayString { get { return Match.Value; } }

        public RegexSearchResult(Match match, CodeModule module, int line)
        {
            Match = match;
            Module = module;
            Selection = new Selection(line, match.Index + 1, line, match.Index + match.Length + 1); // adjust columns for VBE 1-based indexing
        }
    }

    public enum RegexSearchReplaceScope
    {
        Selection,
        CurrentBlock,
        CurrentFile,
        AllOpenedFiles,
        CurrentProject,
        AllOpenProjects
    }
}