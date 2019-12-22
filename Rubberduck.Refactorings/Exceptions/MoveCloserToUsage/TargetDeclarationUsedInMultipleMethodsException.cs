using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.MoveCloserToUsage
{
    public class TargetDeclarationUsedInMultipleMethodsException : InvalidTargetDeclarationException
    {
        public TargetDeclarationUsedInMultipleMethodsException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}