using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Navigation.RegexSearchReplace
{
    public class RegexSearchReplace : IRegexSearchReplace
    {
        private readonly IVBE _vbe;
        private readonly ISelectionService _selectionService;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public RegexSearchReplace(IVBE vbe, ISelectionService selectionService, ISelectedDeclarationProvider selectedDeclarationProvider)
        {
            _vbe = vbe;
            _selectionService = selectionService;
            _selectedDeclarationProvider = selectedDeclarationProvider;
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
                var codeLine = module.GetLines(i, 1);
                var matches = LineMatches(codeLine, searchPattern)
                        .Select(m => new RegexSearchResult(m, module, i));

                results.AddRange(matches);
            }
            return results;
        }

        private IEnumerable<Match> LineMatches(string line, string searchPattern)
        {
            return Regex.Matches(line, searchPattern)
                .OfType<Match>();
        }

        private IEnumerable<Match> LineMatches(string line, int startColumn, int? endColumn, string searchPattern)
        {
            var shortenedLine = endColumn.HasValue 
                ? line.Substring(startColumn - 1, endColumn.Value - startColumn + 1)
                : line.Substring(startColumn - 1);
            return LineMatches(shortenedLine, searchPattern);
        }

        private IEnumerable<RegexSearchResult> GetResultsFromModule(ICodeModule module, string searchPattern, Selection selection)
        {
            var startLine = selection.StartLine > 1
                ? selection.StartLine
                : 1;

            var moduleLines = module.CountOfLines;
            var stopLine = selection.EndLine < moduleLines
                ? selection.EndLine
                : moduleLines;

            if (startLine > stopLine)
            {
                return new List<RegexSearchResult>();
            }

            if (startLine == stopLine)
            {
                return LineMatches(module.GetLines(startLine, 1), selection.StartColumn, null, searchPattern)
                    .Select(m => new RegexSearchResult(m, module, startLine, selection.StartColumn - 1))
                    .ToList();
            }

            var results = new List<RegexSearchResult>();

            var firstLineMatches = LineMatches(module.GetLines(startLine, 1), selection.StartColumn, selection.EndColumn, searchPattern)
                .Select(m => new RegexSearchResult(m, module, startLine));
            results.AddRange(firstLineMatches);

            for (var lineIndex = startLine + 1; lineIndex < stopLine; lineIndex++)
            {
                var codeLine = module.GetLines(lineIndex, 1);
                var matches = LineMatches(codeLine, searchPattern)
                    .Select(m => new RegexSearchResult(m, module, lineIndex));

                results.AddRange(matches);
            }

            var lastLineMatches = LineMatches(module.GetLines(stopLine, 1), 1, selection.EndColumn, searchPattern)
                .Select(m => new RegexSearchResult(m, module, stopLine));
            results.AddRange(lastLineMatches);

            return results;
        }

        private void SetSelection(RegexSearchResult item)
        {
            _selectionService.TrySetActiveSelection(item.Module.QualifiedModuleName, item.Selection);
        }

        private IEnumerable<RegexSearchResult> SearchSelection(string searchPattern)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return new List<RegexSearchResult>();    
                }

                using (var module = pane.CodeModule)
                {
                    return GetResultsFromModule(module, searchPattern, pane.Selection);
                }
            }
        }

        private IEnumerable<RegexSearchResult> SearchCurrentBlock(string searchPattern)
        {
            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return new List<RegexSearchResult>();
            }
            
            var block = _selectedDeclarationProvider
                .SelectedMember(activeSelection.Value)
                ?.Context
                .GetSmallestDescendentContainingSelection<VBAParser.BlockContext>(activeSelection.Value.Selection);

            if (block == null)
            {
                return new List<RegexSearchResult>();
            }

            var blockSelection = block.GetSelection();

            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return new List<RegexSearchResult>();
                }

                using (var module = pane.CodeModule)
                {
                    //FIXME: This is a catastrophe waiting to happen since the module, which will get disposed, is saved on the result.
                    return GetResultsFromModule(module, searchPattern, blockSelection);
                }
            }
        }

        private List<RegexSearchResult> SearchCurrentFile(string searchPattern)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return new List<RegexSearchResult>();
                }

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
