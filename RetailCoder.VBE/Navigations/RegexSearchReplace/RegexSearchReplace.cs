using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

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
                case RegexSearchReplaceScope.Selection:
                    results.AddRange(GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern));
                    results = results.Where(r => _model.Selection.Selection.Contains(r.Selection)).ToList();
                    break;

                case RegexSearchReplaceScope.CurrentBlock:

                    var declarationTypes = new []
                    {
                        DeclarationType.Event,
                        DeclarationType.Function,
                        DeclarationType.Procedure,
                        DeclarationType.PropertyGet,
                        DeclarationType.PropertyLet,
                        DeclarationType.PropertySet
                    };

                    results.AddRange(GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern));
                    dynamic block = _model.ParseResult.Declarations.FindSelection(_model.Selection, declarationTypes).Context.Parent;
                    var selection = new Selection(block.Start.Line, block.Start.Column, block.Stop.Line,
                        block.Stop.Column);
                    results = results.Where(r => selection.Contains(r.Selection)).ToList();
                    break;

                case RegexSearchReplaceScope.CurrentFile:
                    results.AddRange(GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern));
                    break;

                case RegexSearchReplaceScope.AllOpenedFiles:
                    foreach (var codePane in _model.VBE.CodePanes.Cast<CodePane>().Where(codePane => ReferenceEquals(_model.VBE, codePane.VBE)))
                    {
                        results.AddRange(GetResultsFromModule(codePane.CodeModule, searchPattern));
                    }
                    break;

                case RegexSearchReplaceScope.CurrentProject:
                    foreach (dynamic dyn in _model.VBE.ActiveVBProject.VBComponents)
                    {
                        CodeModule module;
                        try
                        {
                            var component = (VBComponent) dyn;
                            module = component.CodeModule;
                        }
                        catch (COMException)
                        {
                            continue;
                        }

                        if (!ReferenceEquals(_model.VBE, module.VBE)) { continue; }

                        results.AddRange(GetResultsFromModule(module, searchPattern));
                    }
                    break;

                case RegexSearchReplaceScope.EntireSolution:
                    foreach (dynamic dyn in _model.VBE.ActiveVBProject.VBComponents)
                    {
                        CodeModule module;
                        try
                        {
                            var component = (VBComponent)dyn;
                            module = component.CodeModule;
                        }
                        catch (COMException)
                        {
                            continue;
                        }

                        if (!ReferenceEquals(_model.VBE, module.VBE)) { continue; }

                        results.AddRange(GetResultsFromModule(module, searchPattern));
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
    }
}