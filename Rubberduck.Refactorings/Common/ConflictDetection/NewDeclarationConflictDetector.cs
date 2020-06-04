using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.Common
{
    public interface INewEntityConflictDetector
    {
        bool HasConflictingName(IConflictDetectionDeclarationProxy proxy, out string nonConflictName, params string[] blackList);
    }

    public interface INewModuleConflictDetector : INewEntityConflictDetector
    {}

    public class NewDeclarationConflictDetector : ConflictDetectorBase, INewEntityConflictDetector, INewModuleConflictDetector
    {
        public NewDeclarationConflictDetector(IDeclarationFinderProvider declarationFinderProvider, 
                                                    IConflictFinderFactory conflictFinderFactory,
                                                    IDeclarationProxyFactory proxyFactory,
                                                    IConflictDetectionSessionData session)
            : base(declarationFinderProvider, conflictFinderFactory, proxyFactory, session)
        {}

        public bool HasConflictingName(IConflictDetectionDeclarationProxy proxy, out string nonConflictName, params string[] blackList)
        {
            var originalIdentifier = proxy.IdentifierName;
            AssignConflictFreeIdentifier(proxy, blackList);

            nonConflictName = proxy.IdentifierName;
            proxy.IdentifierName = originalIdentifier;

            return !nonConflictName.Equals(proxy.IdentifierName);
        }
    }
}
