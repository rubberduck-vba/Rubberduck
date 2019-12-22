using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.Exceptions.ImplementInterface
{
    public class NoImplementsStatementSelectedException : InvalidTargetSelectionException
    {
        public NoImplementsStatementSelectedException(QualifiedSelection targetSelection) 
        :base(targetSelection)
        {}
    }
}
