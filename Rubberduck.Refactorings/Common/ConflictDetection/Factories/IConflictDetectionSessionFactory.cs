using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings
{
    public interface IConflictDetectionSessionFactory
    {
        IConflictDetectionSession Create();
    }

    public class ConflictDetectionSessionFactory : IConflictDetectionSessionFactory
    {
        private readonly IConflictDetectionSessionDataFactory _conflictDetectionSessionDataFactory;

        private IRenameConflictDetection _renamingTools;
        private IRelocateConflictDetection _relocatingTools;
        private INewDeclarationConflictDetection _newDeclarationTools;

        public ConflictDetectionSessionFactory(IDeclarationFinderProvider declarationFinderProvider, IConflictDetectionSessionDataFactory conflictDetectionSessionDataFactory, IConflictFinderFactory conflictFinderFactory)
        {
            _conflictDetectionSessionDataFactory = conflictDetectionSessionDataFactory;

            _renamingTools = new RenameConflictDetection(declarationFinderProvider, conflictFinderFactory);
            _relocatingTools = new RelocateConflictDetection(declarationFinderProvider, conflictFinderFactory);
            _newDeclarationTools = new NewDeclarationConflictDetection(declarationFinderProvider, conflictFinderFactory);
        }

        public IConflictDetectionSession Create()
        {
            return new ConflictDetectionSession(_conflictDetectionSessionDataFactory.Create(), _relocatingTools, _renamingTools, _newDeclarationTools);
        }
    }
}
