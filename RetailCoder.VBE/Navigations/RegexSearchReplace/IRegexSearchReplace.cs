using System.Collections.Generic;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public interface IRegexSearchReplace
    {
        List<RegexSearchResult> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void SearchAndReplace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void SearchAndReplaceAll(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
    }
}