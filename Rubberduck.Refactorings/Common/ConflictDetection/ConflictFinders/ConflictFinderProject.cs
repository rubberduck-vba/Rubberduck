using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts when renaming Project
    /// </summary>
    public class ConflictFinderProject :ConflictFinderBase
    {
        public ConflictFinderProject(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
        : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy projectProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var nameConflicts = new List<Declaration>();

            var matchingProjects = _declarationFinderProvider.DeclarationFinder.Projects
                    .Where(proj => proj.ProjectId != projectProxy.ProjectId
                                        && AreVBAEquivalent(projectProxy.IdentifierName, proj.ProjectName));

            var identifierConflicts = ModuleOrProjectIdentifierConflicts(projectProxy);
            nameConflicts.AddRange(identifierConflicts.Concat(matchingProjects));

            var conflictProxies = CreateProxies(sessionData, nameConflicts);

            var proxyMatches = ModuleOrProjectProxyConflicts(projectProxy, sessionData.RegisteredProxies);

            conflicts.Add(projectProxy, conflictProxies.Concat(proxyMatches).ToList());

            return conflicts.Values.SelectMany(c => c).Any();
        }
    }
}
