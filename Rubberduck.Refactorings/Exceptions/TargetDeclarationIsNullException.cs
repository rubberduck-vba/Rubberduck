using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions
{
    public class TargetDeclarationIsNullException : InvalidTargetDeclarationException
    {
        public TargetDeclarationIsNullException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}