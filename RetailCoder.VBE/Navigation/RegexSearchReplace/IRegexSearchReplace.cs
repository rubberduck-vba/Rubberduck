using System.Collections.Generic;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public interface IRegexSearchReplace
    {
        IEnumerable<RegexSearchResult> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void Replace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        void ReplaceAll(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
    }
}
