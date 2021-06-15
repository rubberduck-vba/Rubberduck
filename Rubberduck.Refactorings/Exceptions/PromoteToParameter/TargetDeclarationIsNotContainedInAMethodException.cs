using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.IntroduceParameter
{
    public class TargetDeclarationIsNotContainedInAMethodException : InvalidTargetDeclarationException
    {
        public TargetDeclarationIsNotContainedInAMethodException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}
