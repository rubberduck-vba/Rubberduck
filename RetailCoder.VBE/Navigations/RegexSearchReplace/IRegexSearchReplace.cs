using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public interface IRegexSearchReplace
    {
        List<QualifiedSelection> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
        List<QualifiedSelection> SearchAndReplace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile);
    }
}