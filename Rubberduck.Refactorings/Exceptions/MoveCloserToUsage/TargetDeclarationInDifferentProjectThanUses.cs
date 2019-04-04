using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.MoveCloserToUsage
{
    public class TargetDeclarationInDifferentProjectThanUses : InvalidTargetDeclarationException
    {
        public TargetDeclarationInDifferentProjectThanUses(Declaration targetDeclaration) 
        :base(targetDeclaration)
        {}
    }
}
