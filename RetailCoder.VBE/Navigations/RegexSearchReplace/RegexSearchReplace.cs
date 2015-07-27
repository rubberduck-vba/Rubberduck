using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly RegexSearchReplaceModel _model;

        public RegexSearchReplace(RegexSearchReplaceModel model)
        {
            _model = model;
        }

        public List<Match> Search(string pattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            if (scope == RegexSearchReplaceScope.CurrentFile)
            {
                var lines = _model.VBE.ActiveCodePane.CodeModule.Lines[0, _model.VBE.ActiveCodePane.CodeModule.CountOfLines];
                return Regex.Matches(lines, pattern).OfType<Match>().ToList();
            }

            return null;
        }

        public void SearchAndReplace(string searchPattern, string replaceValue,
            RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            throw new System.NotImplementedException();
        }
    }
}