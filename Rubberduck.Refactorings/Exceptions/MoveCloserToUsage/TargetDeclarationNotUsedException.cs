using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.MoveCloserToUsage
{
    public class TargetDeclarationNotUsedException : InvalidTargetDeclarationException
    {
        public TargetDeclarationNotUsedException(Declaration targetDeclaration)
        :base(targetDeclaration)
        {}
    }
}