using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Exceptions
{
    public class AffectedModuleIsStaleException : RefactoringException
    {
        public AffectedModuleIsStaleException(QualifiedModuleName staleModule)
        {
            StaleModule = staleModule;
        }

        public QualifiedModuleName StaleModule { get; } 
    }
}