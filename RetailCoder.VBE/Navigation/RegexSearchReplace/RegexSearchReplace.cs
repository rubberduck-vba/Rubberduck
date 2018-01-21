using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly IVBE _vbe;
        private readonly IParseCoordinator _parser;

        public RegexSearchReplace(IVBE vbe, IParseCoordinator parser)
        {
            _vbe = vbe;
            _parser = parser;
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
            return _search.TryGetValue(scope, out var searchFunc) 
                ? searchFunc.Invoke(searchPattern) 
                : new List<RegexSearchResult>();
        }

        public void Replace(string searchPattern, string replaceValue, RegexSearchReplaceScope scope)
        {
            var results = Search(searchPattern, scope).ToList();

            if (results.Count == 0) { return; }

            var result = results[0];

            var originalLine = result.Module.GetLines(results[0].Selection.StartLine, 1);
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

            foreach (var result in results)
            {
                var originalLine = result.Module.GetLines(result.Selection.StartLine, 1);
                var newLine = originalLine.Replace(result.Match.Value, replaceValue);
                result.Module.ReplaceLine(result.Selection.StartLine, newLine);
            }
        }

        private IEnumerable<RegexSearchResult> GetResultsFromModule(ICodeModule module, string searchPattern)
        {
            var results = new List<RegexSearchResult>();

            // VBA uses 1-based indexing
            for (var i = 1; i <= module.CountOfLines; i++)
            {
                var matches =
                    Regex.Matches(module.GetLines(i, 1), searchPattern)
                        .OfType<Match>()
                        .Select(m => new RegexSearchResult(m, module, i));

                results.AddRange(matches);
            }
            return results;
        }

        private void SetSelection(RegexSearchResult item)
        {
            item.Module.CodePane.Selection = item.Selection;
        }

        private List<RegexSearchResult> SearchSelection(string searchPattern)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                using (var module = pane.CodeModule)
                {
                    var results = GetResultsFromModule(module, searchPattern);
                    return results.Where(r => pane.Selection.Contains(r.Selection)).ToList();
                }
            }
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
            using (var pane = _vbe.ActiveCodePane)
            {
                using (var module = pane.CodeModule)
                {
                    var results = GetResultsFromModule(module, searchPattern);

                    using (var moduleParent = module.Parent)
                    {
                        var qualifiedSelection =
                            new QualifiedSelection(new QualifiedModuleName(moduleParent), pane.Selection);
                        dynamic block = state.AllDeclarations.FindTarget(qualifiedSelection, declarationTypes).Context
                            .Parent;
                        var selection = new Selection(block.Start.Line, block.Start.Column, block.Stop.Line,
                            block.Stop.Column);

                        return results.Where(r => selection.Contains(r.Selection)).ToList();
                    }
                }
            }
        }

        private List<RegexSearchResult> SearchCurrentFile(string searchPattern)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                using (var codeModule = pane.CodeModule)
                {
                    return GetResultsFromModule(codeModule, searchPattern).ToList();
                }
            }
        }

        private List<RegexSearchResult> SearchOpenFiles(string searchPattern)
        {
            var results = new List<RegexSearchResult>();
            using (var panes = _vbe.CodePanes)
            {
                foreach (var codePane in panes)
                {
                    try
                    {
                        using (var codeModule = codePane.CodeModule)
                        {
                            results.AddRange(GetResultsFromModule(codeModule, searchPattern));
                        }
                    }
                    finally
                    {
                        codePane.Dispose();
                    }
                }

                return results;
            }
        }

        private List<RegexSearchResult> SearchCurrentProject(string searchPattern)
        {
            var results = new List<RegexSearchResult>();
            using (var project = _vbe.ActiveVBProject)
            {
                using (var components = project.VBComponents)
                {
                    foreach (var component in components)
                    {
                        try
                        {
                            using (var codeModule = component.CodeModule)
                            {
                                results.AddRange(GetResultsFromModule(codeModule, searchPattern));
                            }
                        }
                        finally
                        {
                            component.Dispose();
                        }
                    }
                    return results;
                }
            }
        }

        private List<RegexSearchResult> SearchOpenProjects(string searchPattern)
        {
            var results = new List<RegexSearchResult>();
            using (var projects = _vbe.VBProjects)
            {
                foreach (var project in projects)
                {
                    try
                    {
                        if (project.Protection == ProjectProtection.Locked)
                        {
                            continue;
                        }

                        using (var components = project.VBComponents)
                        {
                            foreach (var component in components)
                            {
                                try
                                {
                                    using (var codeModule = component.CodeModule)
                                    {
                                        results.AddRange(GetResultsFromModule(codeModule, searchPattern));
                                    }
                                }
                                finally
                                {
                                    component.Dispose();
                                }
                            }
                        }
                    }
                    finally
                    {
                        project.Dispose();
                    }
                }
            }
            return results;
        }
    }
}
