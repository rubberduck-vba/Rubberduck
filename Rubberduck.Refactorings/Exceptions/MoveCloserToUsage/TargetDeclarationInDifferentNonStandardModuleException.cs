using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.MoveCloserToUsage
{
    public class TargetDeclarationInDifferentNonStandardModuleException : InvalidTargetDeclarationException
    {
        public TargetDeclarationInDifferentNonStandardModuleException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}