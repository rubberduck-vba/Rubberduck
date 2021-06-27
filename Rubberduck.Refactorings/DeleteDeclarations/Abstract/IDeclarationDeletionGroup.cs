using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationDeletionGroup
    {
        IEnumerable<Declaration> Declarations { get; }
        
        IReadOnlyCollection<IDeclarationDeletionTarget> OrderedFullDeletionTargets { get; }

        IReadOnlyCollection<IDeclarationDeletionTarget> OrderedPartialDeletionTargets { get; }
        
        ParserRuleContext PrecedingNonDeletedContext { set; get; }

        IReadOnlyCollection<IDeclarationDeletionTarget> Targets { get; }
    }
}
