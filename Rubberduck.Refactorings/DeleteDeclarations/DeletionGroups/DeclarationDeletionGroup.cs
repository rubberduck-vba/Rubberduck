using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    internal class DeclarationDeletionGroup : IDeclarationDeletionGroup
    {
        public DeclarationDeletionGroup(IOrderedEnumerable<IDeclarationDeletionTarget> deletionTargets)
        {
            Targets = deletionTargets.ToList();
            OrderedFullDeletionTargets = deletionTargets.Where(dt => dt.IsFullDelete).OrderBy(dt => dt);
            OrderedPartialDeletionTargets = deletionTargets.Where(dt => !dt.IsFullDelete).OrderBy(dt => dt);
        }

        public ParserRuleContext PrecedingNonDeletedContext { set; get; }

        public IReadOnlyCollection<IDeclarationDeletionTarget> Targets { private set; get; }

        public IEnumerable<Declaration> Declarations => Targets.SelectMany(t => t.Declarations);

        public IOrderedEnumerable<IDeclarationDeletionTarget> OrderedFullDeletionTargets { private set; get; }

        public IOrderedEnumerable<IDeclarationDeletionTarget> OrderedPartialDeletionTargets { private set; get; }
    }
}
