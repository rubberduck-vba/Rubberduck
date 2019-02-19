using System.Collections.Generic;

namespace Rubberduck.VBEditor.Utility
{
    public interface ISelectionService
    {
        QualifiedSelection? ActiveSelection();
        ICollection<QualifiedModuleName> OpenModules();
        Selection? Selection(QualifiedModuleName module);
        bool TryActivate(QualifiedModuleName module);
        bool TrySetActiveSelection(QualifiedModuleName module, Selection selection);
        bool TrySetActiveSelection(QualifiedSelection selection);
        bool TrySetSelection(QualifiedModuleName module, Selection selection);
        bool TrySetSelection(QualifiedSelection selection);
    }
}