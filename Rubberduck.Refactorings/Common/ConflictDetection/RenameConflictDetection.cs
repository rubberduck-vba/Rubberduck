using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.Common
{
    public interface IRenameConflictDetection : IConflictDetectionBase
    {
        bool HasRenameConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData);
    }

    public class RenameConflictDetection : ConflictDetectionBase, IRenameConflictDetection
    {
        public RenameConflictDetection(IDeclarationFinderProvider declarationFinderProvider, IConflictFinderFactory conflictFinderFactory)
            : base(declarationFinderProvider, conflictFinderFactory)
        {}

        public bool HasRenameConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData)
        {
            if (AreVBAEquivalent(proxy.Prototype.IdentifierName, proxy.IdentifierName))
            {
                return false;
            }

            var proposedName = proxy.IdentifierName;

            if (TryResolveToConflictFreeIdentifier(proxy, sessionData))
            {
                sessionData.RegisterResolvedProxyIdentifier(proxy);
            }
            return !AreVBAEquivalent(proposedName, proxy.IdentifierName);
        }
    }
}
