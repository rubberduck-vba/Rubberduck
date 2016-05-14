using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly RegexSearchReplaceModel _model;
        private readonly VBE _vbe;
        private readonly IRubberduckParser _parser;
        private readonly ICodePaneWrapperFactory _codePaneFactory;

        public RegexSearchReplace(VBE vbe, IRubberduckParser parser, ICodePaneWrapperFactory codePaneFactory)
        {
            _vbe = vbe;
            _parser = parser;
            _codePaneFactory = codePaneFactory;
            _search = new Dictionary<RegexSearchReplaceScope, Func<string, IEnumerable<RegexSearchResult>>>
            {
                { RegexSearchReplaceScope.Selection, SearchSelection},
                { RegexSearchReplaceScope.CurrentBlock, SearchCurrentBlock},
                { RegexSearchReplaceScope.CurrentFile, SearchCurrentFile},
                { RegexSearchReplaceScope.AllOpenedFiles, SearchOpenFiles},
                { RegexSearchReplaceScope.CurrentProject, SearchCurrentProject},
                { RegexSearchReplaceScope.AllOpenProjects, SearchOpenProjects},
            };
        }

        private readonly IDictionary<RegexSearchReplaceScope,Func<string,IEnumerable<RegexSearchResult>>> _search;

        public IEnumerable<RegexSearchResult> Search(string searchPattern, RegexSearchReplaceScope scope = RegexSearchReplaceScope.CurrentFile)
        {
            Func<string,IEnumerable<RegexSearchResult>> searchFunc;
            if (_search.TryGetValue(scope, out searchFunc))
            {
                return searchFunc.Invoke(searchPattern);
            }
            else
            {
                return new List<RegexSearchResult>();
            }
        }

        public void Replace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            var results = Search(searchPattern, scope).ToList();

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
            var results = Search(searchPattern, scope);

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
            var project = _vbe.ActiveVBProject;
            foreach (var proj in _parser.State.Projects)
            {
                // wtf?
                project = proj;
                break;
            }
            _vbe.SetSelection(project, item.Selection, item.Module.Name, _codePaneFactory);
        }

        private List<RegexSearchResult> SearchSelection(string searchPattern)
        {
            var wrapper = _codePaneFactory.Create(_vbe.ActiveCodePane);
            var results = GetResultsFromModule(_vbe.ActiveCodePane.CodeModule, searchPattern);
            return results.Where(r => wrapper.Selection.Contains(r.Selection)).ToList();
        }

        private List<RegexSearchResult> SearchCurrentBlock(string searchPattern)
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

            var state = _parser.State;
            var results = GetResultsFromModule(_vbe.ActiveCodePane.CodeModule, searchPattern);

            var wrapper = _codePaneFactory.Create(_vbe.ActiveCodePane);
            var qualifiedSelection = new QualifiedSelection(new QualifiedModuleName(wrapper.CodeModule.Parent), wrapper.Selection);
            dynamic block = state.AllDeclarations.FindTarget(qualifiedSelection, declarationTypes).Context.Parent;
            var selection = new Selection(block.Start.Line, block.Start.Column, block.Stop.Line, block.Stop.Column);
            return results.Where(r => selection.Contains(r.Selection)).ToList();
        }

        private List<RegexSearchResult> SearchCurrentFile(string searchPattern)
        {
            return GetResultsFromModule(_vbe.ActiveCodePane.CodeModule, searchPattern).ToList();
        }

        private List<RegexSearchResult> SearchOpenFiles(string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            foreach (var codePane in _vbe.CodePanes.Cast<CodePane>())
            {
                results.AddRange(GetResultsFromModule(codePane.CodeModule, searchPattern));
            }

            return results;
        }

        private List<RegexSearchResult> SearchCurrentProject(string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            foreach (var component in _vbe.ActiveVBProject.VBComponents.Cast<VBComponent>())
            {
                var module = component.CodeModule;
                results.AddRange(GetResultsFromModule(module, searchPattern));
            }

            return results;
        }

        private List<RegexSearchResult> SearchOpenProjects(string searchPattern)
        {
            var results = new List<RegexSearchResult>();
            var modules = _vbe.VBProjects.Cast<VBProject>()
                .Where(project => project.Protection == vbext_ProjectProtection.vbext_pp_none)
                .SelectMany(project => project.VBComponents.Cast<VBComponent>())
                .Select(component => component.CodeModule);

            foreach (var module in modules)
            {
                results.AddRange(GetResultsFromModule(module, searchPattern));
            }

            return results;
        }
    }
}
