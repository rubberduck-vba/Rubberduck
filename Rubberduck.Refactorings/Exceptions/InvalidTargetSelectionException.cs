using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Exceptions
{
    public class InvalidTargetSelectionException : RefactoringException
    {
        public InvalidTargetSelectionException(QualifiedSelection targetSelection)
        {
            TargetSelection = targetSelection;
        }

        public QualifiedSelection TargetSelection { get; }
    }
}