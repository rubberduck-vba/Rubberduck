using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Functions and Subroutines
    /// </summary>
    /// /// See <see cref="ConflictFinderProperties"/> for Properties.
    public class ConflictFinderMembers : ConflictFinderModuleCodeSection
    {
        public ConflictFinderMembers(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy memberProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            if (!IsExistingTargetModule(memberProxy, out var targetModule))
            {
                return false;
            }

            var allMatches = IdentifierMatches(memberProxy, sessionData, out var targetModuleMatches);

            if (TryFindMemberConflictChecksCommon(memberProxy, targetModuleMatches, out var commonChecks))
            {
                conflicts = AddConflicts(conflicts, memberProxy, commonChecks);
            }

            if (NonModuleQualifiedMemberReferenceConflicts(memberProxy, sessionData, allMatches, out var refConflicts))
            {
                conflicts = AddConflicts(conflicts, memberProxy, refConflicts);
            }

            return conflicts.Values.Any();
        }
    }
}
