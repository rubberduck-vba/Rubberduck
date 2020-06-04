using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Functions and Subroutines
    /// </summary>
    /// <seealso cref="ConflictFinderProperties"/>
    public class ConflictFinderMembers : ConflictFinderModuleCodeSection
    {
        public ConflictFinderMembers(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
            : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy memberProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var allMatches = IdentifierMatches(memberProxy, sessionData, out var targetModuleMatches);

            if (TryFindMemberConflictChecksCommon(memberProxy, targetModuleMatches, out var commonChecks))
            {
                conflicts = AddConflicts(conflicts, memberProxy, commonChecks);
            }

            //Check for NonModuleQualifiedReferences to Declarations in other modules
            if (TryFindDeclarationConflictWithOtherNonModuleQualifiedReferences(memberProxy, allMatches, out var conflictRefs3))
            {
                conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, conflictRefs3);
            }

            if (memberProxy.HasStandardModuleParent)
            {
                if (NonModuleQualifiedReferenceConflicts(memberProxy, sessionData, allMatches, out var refConflicts))
                {
                    conflicts = AddConflicts(conflicts, memberProxy, refConflicts);
                }
            }

            return conflicts.Values.Any();
        }
    }
}
