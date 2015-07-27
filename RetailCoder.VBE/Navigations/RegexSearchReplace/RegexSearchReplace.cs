using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly VBProjectParseResult _parseResult;

        public RegexSearchReplace(VBProjectParseResult parseResult)
        {
            _parseResult = parseResult;
        }

        public List<QualifiedSelection> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            throw new System.NotImplementedException();
        }

        public List<QualifiedSelection> SearchAndReplace(string searchPattern, string replaceValue,
            RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            throw new System.NotImplementedException();
        }
    }
}