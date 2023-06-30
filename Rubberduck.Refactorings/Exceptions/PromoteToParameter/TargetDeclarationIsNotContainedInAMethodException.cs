using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.PromoteToParameter
{
    public class TargetDeclarationIsNotContainedInAMethodException : InvalidTargetDeclarationException
    {
        public TargetDeclarationIsNotContainedInAMethodException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}
