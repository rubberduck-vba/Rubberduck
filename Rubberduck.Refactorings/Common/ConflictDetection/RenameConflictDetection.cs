using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.Common
{
    public interface IRenameConflictDetection : IConflictDetectionBase
    {}

    public class RenameConflictDetection : ConflictDetectionBase, IRenameConflictDetection
    {
        public RenameConflictDetection(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory)
            : base(declarationFinderProvider, conflictFinderFactory)
        {}

        public override bool HasConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            if (AreVBAEquivalent(proxy.Prototype.IdentifierName, proxy.IdentifierName))
            {
                return false;
            }
            return !CanResolveToConflictFreeIdentifier(proxy, sessionData);
        }
    }
}
