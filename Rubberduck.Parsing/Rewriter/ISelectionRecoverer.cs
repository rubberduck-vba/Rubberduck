using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Rewriter
{
    public interface ISelectionRecoverer
    {
        void SaveSelections(IEnumerable<QualifiedModuleName> modules);
        void AdjustSavedSelection(QualifiedModuleName module, Selection selectionOffset);
        void ReplaceSavedSelection(QualifiedModuleName module, Selection replacementSelection);
        void RecoverSavedSelections();
        void RecoverSavedSelectionsOnNextParse();

        void SaveActiveCodePane();
        void RecoverActiveCodePane();
        void RecoverActiveCodePaneOnNextParse();

        void SaveOpenState(IEnumerable<QualifiedModuleName> modules);
        void RecoverOpenState();
        void RecoverOpenStateOnNextParse();
    }
}