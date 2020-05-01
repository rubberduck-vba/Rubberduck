using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Modules
    /// </summary>
    public class ConflictFinderModule : ConflictFinderBase
    {
        public ConflictFinderModule(IDeclarationFinderProvider declarationFinderProvider)
        : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy moduleProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var moduleNameConflicts = MatchingIdentifierDeclarations(moduleProxy,
                                                    DeclarationType.UserDefinedType,
                                                    DeclarationType.Enumeration,
                                                    DeclarationType.Module)
                                                    .ToList();

            if (AreVBAEquivalent(moduleProxy.ProjectName, moduleProxy.IdentifierName))
            {
                moduleNameConflicts.Add(moduleProxy.ParentDeclaration);
            }

            var conflictProxies = CreateProxies(sessionData, moduleNameConflicts);
            conflicts.Add(moduleProxy, conflictProxies.ToList());

            return conflicts.Values.SelectMany(lst => lst).Any();
        }
    }
}
