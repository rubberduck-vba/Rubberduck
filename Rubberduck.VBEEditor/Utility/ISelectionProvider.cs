using System.Collections.Generic;

namespace Rubberduck.VBEditor.Utility
{
    public interface ISelectionProvider
    {
        QualifiedSelection? ActiveSelection();
        ICollection<QualifiedModuleName> OpenModules();
        Selection? Selection(QualifiedModuleName module);
    }
}