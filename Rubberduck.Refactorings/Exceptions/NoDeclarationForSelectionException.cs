using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Exceptions
{
    public class NoDeclarationForSelectionException : InvalidTargetSelectionException
    {
        public NoDeclarationForSelectionException(QualifiedSelection targetSelection)
        :base(targetSelection)
        {}
    }
}