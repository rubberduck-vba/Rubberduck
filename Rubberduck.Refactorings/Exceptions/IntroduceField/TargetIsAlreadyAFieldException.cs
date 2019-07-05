using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.IntroduceField
{
    public class TargetIsAlreadyAFieldException : InvalidTargetDeclarationException
    {
        public TargetIsAlreadyAFieldException(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}
