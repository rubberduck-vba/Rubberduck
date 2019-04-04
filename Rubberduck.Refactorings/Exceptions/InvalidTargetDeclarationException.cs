using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions
{
    public class InvalidTargetDeclarationException : RefactoringException
    {
        public InvalidTargetDeclarationException(Declaration targetDeclaration)
        {
            TargetDeclaration = targetDeclaration;
        }

        public Declaration TargetDeclaration { get; }
    }
}