using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Exceptions
{
    public class InvalidTargetSelectionException : RefactoringException
    {
        public InvalidTargetSelectionException(QualifiedSelection targetSelection)
        {
            TargetSelection = targetSelection;
        }

        public InvalidTargetSelectionException(QualifiedSelection targetSelection, string message)
        {
            TargetSelection = targetSelection;
            Message = message;
        }

        public QualifiedSelection TargetSelection { get; }
        public override string Message { get; }
    }
}