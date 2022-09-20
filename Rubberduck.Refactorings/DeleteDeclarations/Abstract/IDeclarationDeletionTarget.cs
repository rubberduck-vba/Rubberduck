using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using System.Collections.Generic;

namespace Rubberduck.Refactorings
{
    public interface IModuleElementDeletionTarget : IDeclarationDeletionTarget
    { }

    public interface IPropertyDeletionTarget
    {
        bool IsGroupedWithRelatedProperties();
    }

    public interface IEnumMemberDeletionTarget : IDeclarationDeletionTarget
    {}

    public interface IUdtMemberDeletionTarget : IDeclarationDeletionTarget
    {}

    public interface ILocalScopeDeletionTarget : IDeclarationDeletionTarget
    {
        ILocalScopeDeletionTarget AssociatedLabelToDelete { get; }

        bool IsLabel(out ILabelDeletionTarget labelTarget);

        ParserRuleContext ScopingContext { get; }
        bool HasSameLogicalLineLabel(out VBAParser.StatementLabelDefinitionContext labelContext);

        void SetupToDeleteAssociatedLabel(ILabelDeletionTarget label);

    }

    public interface ILabelDeletionTarget
    {
        bool HasSameLogicalLineListContext(out ParserRuleContext varOrConst);

        bool HasFollowingMainBlockStatementContext(out VBAParser.MainBlockStmtContext mainBlockStmtContext);

        bool ReplaceLabelWithWhitespace { set; get; }
    }

    public interface IDeclarationDeletionTarget
    {
        bool IsFullDelete { get; }

        void AddTargets(IEnumerable<Declaration> targets);

        /// <summary>
        /// TargetProxy is the target for DeclarationTypes that are not declared
        /// in a list, or the first Declaration in a Declaration List.
        /// </summary>
        Declaration TargetProxy { get; }

        IReadOnlyCollection<Declaration> Declarations { get; }

        IModuleRewriter Rewriter { get; }

        IReadOnlyList<Declaration> AllDeclarationsInListContext { get; }

        IEnumerable<Declaration> RetainedDeclarations { get; }

        VBAParser.EndOfStatementContext PrecedingEOSContext { set; get; }

        VBAParser.EndOfStatementContext TargetEOSContext { get; }

        VBAParser.EndOfStatementContext EOSContextToReplace { get; }

        ParserRuleContext DeleteContext { get; }

        ParserRuleContext ListContext { get; }

        ParserRuleContext TargetContext { get; }

        VBAParser.CommentContext GetDeclarationLogicalLineCommentContext();

        string BuildEOSReplacementContent();

        string ModifiedTargetEOSContent { get; }

        bool DeletionIncludesEOSContext { get; }
    }
}
