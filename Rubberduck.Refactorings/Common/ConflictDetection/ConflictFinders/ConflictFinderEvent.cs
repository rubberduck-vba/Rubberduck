﻿using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public class ConflictFinderEvent : ConflictFinderModuleDeclarationSection
    {
        public ConflictFinderEvent(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var allMatches = IdentifierMatches(proxy, sessionData, out _);

            conflicts = AddConflicts(conflicts, proxy, allMatches.Where(d => d.DeclarationType.HasFlag(DeclarationType.Event)));
            return conflicts.Values.Any();
        }
    }
}
