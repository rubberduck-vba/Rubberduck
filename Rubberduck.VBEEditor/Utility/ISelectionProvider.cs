using System.Collections.Generic;

namespace Rubberduck.VBEditor.Utility
{
    public interface ISelectionProvider
    {
        /// <summary>
        /// Gets the QualifiedModuleName for the component that is currently selected in the Project Explorer.
        /// </summary>
        QualifiedModuleName ProjectExplorerSelection();
        QualifiedSelection? ActiveSelection();
        ICollection<QualifiedModuleName> OpenModules();
        Selection? Selection(QualifiedModuleName module);
    }
}