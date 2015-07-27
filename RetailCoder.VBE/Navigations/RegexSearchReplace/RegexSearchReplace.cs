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

        public List<RegexSearchResult> Search(string searchPattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            var results = new List<RegexSearchResult>();

            if (scope == RegexSearchReplaceScope.CurrentFile)
            {
                var module = _model.VBE.ActiveCodePane.CodeModule;

                for (var i = 0; i < module.CountOfLines; i++)
                {
                    results.AddRange(Regex.Matches(module.Lines[i, 1], searchPattern).OfType<Match>().Select(m => new RegexSearchResult(m, i)));
                }
            }

            return results;
        }

        public void SearchAndReplace(string searchPattern, string replaceValue,
            RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            if (scope == RegexSearchReplaceScope.CurrentFile)
            {
                var module = _model.VBE.ActiveCodePane.CodeModule;
                var results = Search(searchPattern);

                foreach (var result in results)
                {
                    var originalLine = module.Lines[result.Selection.StartLine, 1];
                    var newLine = originalLine.Replace(result.Match.Value, replaceValue);
                    module.ReplaceLine(result.Selection.StartLine, newLine);
                }
            }
        }
    }
}