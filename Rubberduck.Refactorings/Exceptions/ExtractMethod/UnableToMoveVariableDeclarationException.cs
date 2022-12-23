using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.ExtractMethod
{
    public class UnableToMoveVariableDeclarationException : RefactoringException
    {
        public override string Message { get; }
        public UnableToMoveVariableDeclarationException(Declaration problemDeclaration)
        {
            Message = problemDeclaration.Context.GetText();
        }
    }
}

