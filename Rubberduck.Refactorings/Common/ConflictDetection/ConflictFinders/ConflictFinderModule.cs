using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Modules
    /// </summary>
    public class ConflictFinderModule : ConflictFinderBase
    {
        public ConflictFinderModule(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
        : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy moduleProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var moduleNameConflicts = new List<Declaration>();
            if (ModuleIdentifierMatchesProjectName(moduleProxy, out var project))
            {
                moduleNameConflicts.Add(project);
            }

            var moduleIdentifierConflicts = ModuleOrProjectIdentifierConflicts(moduleProxy);
            moduleNameConflicts.AddRange(moduleIdentifierConflicts);

            var conflictProxies = CreateProxies(sessionData, moduleNameConflicts).ToList();

            var proxyMatches = ModuleOrProjectProxyConflicts(moduleProxy, sessionData.RegisteredProxies);

            conflictProxies.AddRange(proxyMatches);

            conflicts.Add(moduleProxy, conflictProxies);

            return conflicts.Values.SelectMany(c => c).Any();
        }

        private bool ModuleIdentifierMatchesProjectName(IConflictDetectionDeclarationProxy moduleProxy, out Declaration project)
        {
            project = _declarationFinderProvider.DeclarationFinder.Projects
                    .FirstOrDefault(proj => proj.ProjectId == moduleProxy.ProjectId
                                        && AreVBAEquivalent(moduleProxy.IdentifierName, proj.ProjectName));

            return project != null;
        }
    }
}
