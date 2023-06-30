using Antlr4.Runtime;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.DeleteDeclarations;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings
{
    public interface IDeclarationDeletionTargetFactory
    {
        IDeclarationDeletionTarget Create(Declaration declaration, IRewriteSession rewriteSession);

        IEnumerable<IDeclarationDeletionTarget> CreateMany(IEnumerable<Declaration> declarations, IRewriteSession rewriteSession);
    }

    public interface IDeclarationDeletionGroupFactory
    {
        IDeclarationDeletionGroup Create(IOrderedEnumerable<IDeclarationDeletionTarget> deletionTargets);
    }

    public interface IDeclarationDeletionGroupsGeneratorFactory
    {
        IDeclarationDeletionGroupsGenerator Create();
    }
}
