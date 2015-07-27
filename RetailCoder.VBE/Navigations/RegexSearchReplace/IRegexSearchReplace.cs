using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public interface IRegexSearchReplace
    {
        List<RegexSearchResult> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void SearchAndReplace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
    }
}