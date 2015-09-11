using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly RegexSearchReplaceModel _model;
        private readonly ICodePaneWrapperFactory _codePaneFactory;

        public RegexSearchReplace (RegexSearchReplaceModel model, ICodePaneWrapperFactory codePaneFactory)
        {
            _model = model;
            _codePaneFactory = codePaneFactory;
        }

        public List<RegexSearchResult> Find(string searchPattern, RegexSearchReplaceScope scope)
        {
            switch (scope)
            {
                case RegexSearchReplaceScope.Selection:
                    return GetResults_Selection(searchPattern);

                case RegexSearchReplaceScope.CurrentBlock:
                    return GetResults_CurrentBlock(searchPattern);

                case RegexSearchReplaceScope.CurrentFile:
                    return GetResults_CurrentFile(searchPattern);

                case RegexSearchReplaceScope.AllOpenedFiles:
                    return GetResults_AllOpenedFiles(searchPattern);

                case RegexSearchReplaceScope.CurrentProject:
                    return GetResults_CurrentProject(searchPattern);

                case RegexSearchReplaceScope.AllOpenProjects:
                    return GetResults_AllOpenProjects(searchPattern);

                default:
                    return new List<RegexSearchResult>();
            }
        }

        public void Replace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            var results = Find(searchPattern, scope);

            if (results.Count == 0) { return; }

            RegexSearchResult result = results[0];

            string originalLine = result.Module.Lines[results[0].Selection.StartLine, 1];
            var newLine = originalLine.Replace(result.Match.Value, replaceValue);
            result.Module.ReplaceLine(result.Selection.StartLine, newLine);

            if (results.Count >= 2)
            {
                SetSelection(results[1]);
            }
        }

        public void ReplaceAll(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            var results = Find(searchPattern, scope);

            foreach (RegexSearchResult result in results)
            {
                string originalLine = result.Module.Lines[result.Selection.StartLine, 1];
                var newLine = originalLine.Replace(result.Match.Value, replaceValue);
                result.Module.ReplaceLine(result.Selection.StartLine, newLine);
            }
        }

        private IEnumerable<RegexSearchResult> GetResultsFromModule(CodeModule module, string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            // VBA uses 1-based indexing
            for (var i = 1; i <= module.CountOfLines; i++)
            {
                var matches =
                    Regex.Matches(module.Lines[i, 1], searchPattern)
                        .OfType<Match>()
                        .Select(m => new RegexSearchResult(m, module, i));

                results.AddRange(matches);
            }
            return results;
        }

        private void SetSelection(RegexSearchResult item)
        {
            var project = item.Module.VBE.ActiveVBProject;
            foreach (var proj in item.Module.VBE.VBProjects.Cast<VBProject>().Where(proj => proj.VBComponents.Cast<VBComponent>().Any(v => v.CodeModule == item.Module)))
            {
                project = proj;
                break;
            }
            _model.VBE.SetSelection(project, item.Selection, item.Module.Name, _codePaneFactory);
        }

        private List<RegexSearchResult> GetResults_Selection(string searchPattern)
        {
            IEnumerable<RegexSearchResult> results = GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern);
            return results.Where(r => _model.Selection.Selection.Contains(r.Selection)).ToList();
        }

        private List<RegexSearchResult> GetResults_CurrentBlock(string searchPattern)
        {
            var declarationTypes = new[]
                    {
                        DeclarationType.Event,
                        DeclarationType.Function,
                        DeclarationType.Procedure,
                        DeclarationType.PropertyGet,
                        DeclarationType.PropertyLet,
                        DeclarationType.PropertySet
                    };

            IEnumerable<RegexSearchResult> results = GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern);
            dynamic block = _model.ParseResult.Declarations.FindSelection(_model.Selection, declarationTypes).Context.Parent;
            var selection = new Selection(block.Start.Line, block.Start.Column, block.Stop.Line,
                block.Stop.Column);
            return results.Where(r => selection.Contains(r.Selection)).ToList();
        }

        private List<RegexSearchResult> GetResults_CurrentFile(string searchPattern)
        {
            return GetResultsFromModule(_model.VBE.ActiveCodePane.CodeModule, searchPattern).ToList();
        }

        private List<RegexSearchResult> GetResults_AllOpenedFiles(string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            foreach (var codePane in _model.VBE.CodePanes.Cast<CodePane>().Where(codePane => ReferenceEquals(_model.VBE, codePane.VBE)))
            {
                results.AddRange(GetResultsFromModule(codePane.CodeModule, searchPattern));
            }

            return results;
        }

        private List<RegexSearchResult> GetResults_CurrentProject(string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            foreach (var component in _model.VBE.ActiveVBProject.VBComponents.Cast<VBComponent>())
            {
                var module = component.CodeModule;

                if (!ReferenceEquals(_model.VBE.ActiveVBProject, module.VBE.ActiveVBProject)) { continue; }
                results.AddRange(GetResultsFromModule(module, searchPattern));
            }

            return results;
        }

        private List<RegexSearchResult> GetResults_AllOpenProjects(string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            foreach (VBProject project in _model.VBE.VBProjects)
            {
                foreach (var component in project.VBComponents.Cast<VBComponent>())
                {
                    var module = component.CodeModule;

                    if (!ReferenceEquals(_model.VBE, module.VBE))
                    {
                        continue;
                    }
                    results.AddRange(GetResultsFromModule(module, searchPattern));
                }
            }

            return results;
        }
    }
}