using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.MoveCloserToUsage
{
    public class TargetDeclarationNonPrivateInNonStandardModule : InvalidTargetDeclarationException
    {
        public TargetDeclarationNonPrivateInNonStandardModule(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}