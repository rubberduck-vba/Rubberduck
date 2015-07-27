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

        public List<RegexSearchResult> Search(string searchPattern, RegexSearchReplaceScope scope)
        {
            var results = new List<RegexSearchResult>();

            if (scope == RegexSearchReplaceScope.CurrentFile)
            {
                var module = _model.VBE.ActiveCodePane.CodeModule;

                for (var i = 1; i <= module.CountOfLines; i++)
                {
                    var matches =
                        Regex.Matches(module.Lines[i, 1], searchPattern)
                            .OfType<Match>()
                            .Select(m => new RegexSearchResult(m, i));

                    if (matches.Any())
                    {
                        results.AddRange(matches);
                    }
                }
            }

            return results;
        }

        public void SearchAndReplace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            if (scope == RegexSearchReplaceScope.CurrentFile)
            {
                var module = _model.VBE.ActiveCodePane.CodeModule;
                var results = Search(searchPattern, scope);

                if (results.Count > 0)
                {
                    var originalLine = module.Lines[results[0].Selection.StartLine, 1];
                    var newLine = originalLine.Replace(results[0].Match.Value, replaceValue);
                    module.ReplaceLine(results[0].Selection.StartLine, newLine);
                }
            }
        }

        public void SearchAndReplaceAll(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            if (scope == RegexSearchReplaceScope.CurrentFile)
            {
                var module = _model.VBE.ActiveCodePane.CodeModule;
                var results = Search(searchPattern, scope);

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