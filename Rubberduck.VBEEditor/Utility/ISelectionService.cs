namespace Rubberduck.VBEditor.Utility
{
    public interface ISelectionService
    {
        QualifiedSelection? ActiveSelection();
        Selection? Selection(QualifiedModuleName module);
        bool TryActivate(QualifiedModuleName module);
        bool TrySetActiveSelection(QualifiedSelection selection);
        bool TrySetSelection(QualifiedModuleName module, Selection selection);
        bool TrySetSelection(QualifiedSelection selection);
    }
}