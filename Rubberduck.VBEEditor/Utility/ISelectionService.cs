namespace Rubberduck.VBEditor.Utility
{
    public interface ISelectionService : ISelectionProvider
    {
        bool TryActivate(QualifiedModuleName module);
        bool TrySetActiveSelection(QualifiedModuleName module, Selection selection);
        bool TrySetActiveSelection(QualifiedSelection selection);
        bool TrySetSelection(QualifiedModuleName module, Selection selection);
        bool TrySetSelection(QualifiedSelection selection);
    }
}