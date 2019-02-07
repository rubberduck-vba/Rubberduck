using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Parsing.Rewriter
{
    public class SelectionRecoverer : ISelectionRecoverer
    {
        private readonly ISelectionService _selectionService;
        private readonly IParseManager _parseManager;

        private readonly IDictionary<QualifiedModuleName, Selection> _savedSelections = new Dictionary<QualifiedModuleName, Selection>();
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
    }
}