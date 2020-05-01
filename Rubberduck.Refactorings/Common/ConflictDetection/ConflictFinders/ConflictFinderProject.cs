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
    /// Finds identifier conflicts when renaming Project
    /// </summary>
    public class ConflictFinderProject :ConflictFinderBase
    {
        public ConflictFinderProject(IDeclarationFinderProvider declarationFinderProvider)
        : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy projectProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var conflictingDeclarations = MatchingIdentifierDeclarations(projectProxy,
                                                    DeclarationType.UserDefinedType,
                                                    DeclarationType.Enumeration)
                                                    .ToList();

            var conflictingProxies = CreateProxies(sessionData, conflictingDeclarations);

            conflicts.Add(projectProxy, conflictingProxies.ToList());

            return conflicts.Values.SelectMany(lst => lst).Any();
        }
    }
}
