using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;

namespace Rubberduck.Refactorings
{
    public interface IConflictSessionFactory
    {
        IConflictSession Create();
    }

    public class ConflictSessionFactory : IConflictSessionFactory
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IConflictDetectorFactory _detectorFactory;
        private readonly IDeclarationProxyFactory _proxyFactory;

        public ConflictSessionFactory(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory, IConflictDetectorFactory detectorFactory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _proxyFactory = proxyFactory;
            _detectorFactory = detectorFactory;
        }

        public IConflictSession Create()
        {
            return new ConflictDetectionSession(_declarationFinderProvider, _proxyFactory, _detectorFactory, new ConflictDetectionSessionData());
        }
    }
}
