using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.Exceptions.MoveCloserToUsage
{
    public class TargetDeclarationConflictsWithPreexistingDeclaration : InvalidTargetDeclarationException
    {
        public TargetDeclarationConflictsWithPreexistingDeclaration(Declaration targetDeclaration, Declaration conflictingDeclaration)
        :base(targetDeclaration)
        {
            ConflictingDeclaration = conflictingDeclaration;
        }

        public Declaration ConflictingDeclaration { get; }
    }
}