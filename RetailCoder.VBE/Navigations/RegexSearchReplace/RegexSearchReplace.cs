using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Navigations.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly RegexSearchReplaceModel _model;

        public RegexSearchReplace(RegexSearchReplaceModel model)
        {
            _model = model;
        }

        public List<RegexSearchResult> Find(string searchPattern, RegexSearchReplaceScope scope)
        {
            var results = new List<RegexSearchResult>();

            switch (scope)
            {
                case RegexSearchReplaceScope.CurrentFile:
                    results.AddRange(GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern));
                    break;

                case RegexSearchReplaceScope.AllOpenedFiles:
                    foreach (var codePane in _model.VBE.CodePanes.Cast<CodePane>().Where(codePane => ReferenceEquals(_model.VBE, codePane.VBE)))
                    {
                        results.AddRange(GetResultsFromModule(codePane.CodeModule, searchPattern));
                    }
                    break;
            }

            return results;
        }

        public void Replace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            var results = Find(searchPattern, scope);

            if (results.Count <= 0) { return; }

            var originalLine = results[0].Module.Lines[results[0].Selection.StartLine, 1];
            var newLine = originalLine.Replace(results[0].Match.Value, replaceValue);
            results[0].Module.ReplaceLine(results[0].Selection.StartLine, newLine);
        }

        public void ReplaceAll(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            var results = Find(searchPattern, scope);

            foreach (var result in results)
            {
                var originalLine = result.Module.Lines[result.Selection.StartLine, 1];
                var newLine = originalLine.Replace(result.Match.Value, replaceValue);
                result.Module.ReplaceLine(result.Selection.StartLine, newLine);
            }
        }

        private IEnumerable<RegexSearchResult> GetResultsFromModule(CodeModule module, string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            for (var i = 1; i <= module.CountOfLines; i++)
            {
                var matches =
                    Regex.Matches(module.Lines[i, 1], searchPattern)
                        .OfType<Match>()
                        .Select(m => new RegexSearchResult(m, module, i)).ToList();

                if (matches.Any())
                {
                    results.AddRange(matches);
                }
            }
            return results;
        }

        /*private void ShowResultsToolwindow(IEnumerable<Declaration> implementations, string name)
        {
            // throws a COMException if toolwindow was already closed
            var window = new SimpleListControl(string.Format(RubberduckUI.RegexSearchReplace_Caption, name));
            var presenter = new ImplementationsListDockablePresenter(_vbe, _addIn, window, implementations, _codePaneFactory);
            presenter.Show();
        }*/
    }
}