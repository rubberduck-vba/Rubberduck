using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.Rewriter
{
    public sealed class SelectionRecoverer : ISelectionRecoverer, IDisposable
    {
        private readonly ISelectionService _selectionService;
        private readonly IParseManager _parseManager;

        private readonly IDictionary<QualifiedModuleName, Selection> _savedSelections = new Dictionary<QualifiedModuleName, Selection>();
        private readonly HashSet<QualifiedModuleName> _savedOpenModules = new HashSet<QualifiedModuleName>();
        private QualifiedModuleName? _savedActiveModule = null;
        private const ParserState ParserStateOnWhichToTriggerRecovery = ParserState.LoadingReference;

        public SelectionRecoverer(ISelectionService selectionService, IParseManager parseManager)
        {
            _selectionService = selectionService;
            _parseManager = parseManager;
        }


        public void SaveSelections(IEnumerable<QualifiedModuleName> modules)
        {
            _savedSelections.Clear();
            foreach (var module in modules.Distinct())
            {
                var selection = _selectionService.Selection(module);
                if (selection.HasValue)
                {
                    _savedSelections.Add(module, selection.Value);
                }
            }
        }

        public void AdjustSavedSelection(QualifiedModuleName module, Selection selectionOffset)
        {
            if (_savedSelections.TryGetValue(module, out var savedSelection))
            {
                _savedSelections[module] = savedSelection.Offset(selectionOffset);
            }
        }

        public void ReplaceSavedSelection(QualifiedModuleName module, Selection replacementSelection)
        {
            _savedSelections[module] = replacementSelection;
        }

        public void RecoverSavedSelections()
        {
            foreach (var (module, selection) in _savedSelections)
            {
                _selectionService.TrySetSelection(module, selection);
            }
            _savedSelections.Clear();
        }

        public void RecoverSavedSelectionsOnNextParse()
        {
            _parseManager.StateChanged += ExecuteSelectionRecovery;
        }

        private void ExecuteSelectionRecovery(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserStateOnWhichToTriggerRecovery)
            {
                _parseManager.StateChanged -= ExecuteSelectionRecovery;
                RecoverSavedSelections();
            }
        }


        public void SaveActiveCodePane()
        {
            _savedActiveModule = _selectionService.ActiveSelection()?.QualifiedName;
        }

        public void RecoverActiveCodePane()
        {
            if (_savedActiveModule.HasValue)
            {
                _selectionService.TryActivate(_savedActiveModule.Value);
                _savedActiveModule = null;
            }
        }

        public void RecoverActiveCodePaneOnNextParse()
        {
            _parseManager.StateChanged += ExecuteActiveCodePaneRecovery;
        }

        private void ExecuteActiveCodePaneRecovery(object sender, ParserStateEventArgs e)
        {
            if (e.State == ParserStateOnWhichToTriggerRecovery)
            {
                _parseManager.StateChanged -= ExecuteActiveCodePaneRecovery;
                RecoverActiveCodePane();
            }
        }

        public void SaveOpenState(IEnumerable<QualifiedModuleName> modules)
        {
            _savedOpenModules.Clear();
            var openModules = _selectionService.OpenModules();
            _savedOpenModules.UnionWith(modules.Where(module => openModules.Contains(module)));
        }

        public void RecoverOpenState()
        {
            foreach (var module in _savedOpenModules)
            {
                _selectionService.TryActivate(module);
            }
            _savedOpenModules.Clear();
        }

        public void Dispose()
        {
            _parseManager.StateChanged -= ExecuteSelectionRecovery;
            _parseManager.StateChanged -= ExecuteActiveCodePaneRecovery;
        }
    }
}