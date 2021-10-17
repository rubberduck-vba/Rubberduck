using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationDeletionGroup
    {
        IEnumerable<Declaration> Declarations { get; }

        IOrderedEnumerable<IDeclarationDeletionTarget> OrderedFullDeletionTargets { get; }

        IOrderedEnumerable<IDeclarationDeletionTarget> OrderedPartialDeletionTargets { get; }
        
        ParserRuleContext PrecedingNonDeletedContext { set; get; }

        IReadOnlyCollection<IDeclarationDeletionTarget> Targets { get; }
    }
}
