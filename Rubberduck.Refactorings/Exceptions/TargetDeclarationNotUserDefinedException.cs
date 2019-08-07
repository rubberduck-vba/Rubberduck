using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions
{
    public class TargetDeclarationNotUserDefinedException : InvalidTargetDeclarationException
    {
        public TargetDeclarationNotUserDefinedException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}