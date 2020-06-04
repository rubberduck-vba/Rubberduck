using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;

namespace Rubberduck.Refactorings
{
    public interface IConflictDetectorFactory
    {
        IRenameConflictDetector CreateRenameConflictDetector(IConflictDetectionSessionData sessionData);
        IRelocateConflictDetector CreateRelocateConflictDetector(IConflictDetectionSessionData sessionData);
        INewModuleConflictDetector CreateNewModuleConflictDetector(IConflictDetectionSessionData sessionData);
        INewEntityConflictDetector CreateNewEntityConflictDetector(IConflictDetectionSessionData sessionData);
    }

    public class ConflictDetectorFactory : IConflictDetectorFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IConflictFinderFactory _conflictFinderFactory;
        private readonly IDeclarationProxyFactory _proxyFactory;

        public ConflictDetectorFactory(IDeclarationFinderProvider declarationFinderProvider,
                                        IConflictFinderFactory conflictFinderFactory,
                                        IDeclarationProxyFactory proxyFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _conflictFinderFactory = conflictFinderFactory;
            _proxyFactory = proxyFactory;
        }

        public IRenameConflictDetector CreateRenameConflictDetector(IConflictDetectionSessionData sessionData)
        {
            return new RenameConflictDetector(_declarationFinderProvider, _conflictFinderFactory, _proxyFactory, sessionData);
        }

        public IRelocateConflictDetector CreateRelocateConflictDetector(IConflictDetectionSessionData sessionData)
        {
            return new RelocateConflictDetector(_declarationFinderProvider, _conflictFinderFactory, _proxyFactory, sessionData);
        }

        public INewModuleConflictDetector CreateNewModuleConflictDetector(IConflictDetectionSessionData sessionData)
        {
            return new NewDeclarationConflictDetector(_declarationFinderProvider, _conflictFinderFactory, _proxyFactory, sessionData) as INewModuleConflictDetector;
        }

        public INewEntityConflictDetector CreateNewEntityConflictDetector(IConflictDetectionSessionData sessionData)
        {
            return new NewDeclarationConflictDetector(_declarationFinderProvider, _conflictFinderFactory, _proxyFactory, sessionData);
        }
    }
}
