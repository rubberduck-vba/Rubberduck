using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Events
    /// </summary>
    public class ConflictFinderEvent : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderEvent(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
            : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var allMatches = IdentifierMatches(proxy, sessionData, out _);

            conflicts = AddConflicts(conflicts, proxy, allMatches.Where(d => d.DeclarationType.HasFlag(DeclarationType.Event) && d.QualifiedModuleName == proxy.QualifiedModuleName));
            return conflicts.Values.Any();
        }
    }
}
