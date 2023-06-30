using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class ProcedureLocalDeletionTarget<T> : DeclarationDeletionTargetBase, ILocalScopeDeletionTarget where T : ParserRuleContext
    {
        public ProcedureLocalDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
            : base(declarationFinderProvider, target, rewriter)
        {
            ListContext = target.Context.GetAncestor<T>();

            TargetContext = target.Context.GetAncestor<VBAParser.BlockStmtContext>();

            //If there is a label on the declaration's line, delete just the declaration's context
            DeleteContext = TargetContext.TryGetChildContext<VBAParser.StatementLabelDefinitionContext>(out _)
                ? TargetContext.GetChild<VBAParser.MainBlockStmtContext>()
                : target.Context.GetAncestor<VBAParser.BlockStmtContext>() as ParserRuleContext;

            //The preceding EOS Context cannot be determined directly from the target.  It depends upon 
            //what other targets are deleted adjacent to the target.
            PrecedingEOSContext = null;

            TargetEOSContext = DeleteContext.GetFollowingEndOfStatementContext();

            switch (TargetContext.Parent.Parent)
            {
                case VBAParser.ForNextStmtContext forNext:
                    ScopingContext = forNext.GetChild<VBAParser.UnterminatedBlockContext>();
                    break;
                case VBAParser.ForEachStmtContext forEach:
                    ScopingContext = forEach.GetChild<VBAParser.UnterminatedBlockContext>();
                    break;
                default:
                    ScopingContext = TargetContext.Parent as ParserRuleContext;
                    break;
            }

            //Initializes for the default use case where a Label exists on the same logical line as a Variable/Constant to be  
            //deleted.  The VBAParser.EndOfStatementContext (of the Variable/Const) is retained to provide spacing
            //and indentation for the next BlockStatementContext.  If the Label is also to be deleted, this flag is modified.
            DeletionIncludesEOSContext = !HasSameLogicalLineLabel(out _);
        }

        public override bool IsFullDelete
            => AllDeclarationsInListContext.Intersect(Targets).Count() == AllDeclarationsInListContext.Count;

        public ILocalScopeDeletionTarget AssociatedLabelToDelete { private set; get; }

        public void SetupToDeleteAssociatedLabel(ILabelDeletionTarget label)
        {
            AssociatedLabelToDelete = label as ILocalScopeDeletionTarget;
            DeletionIncludesEOSContext = label != null;
        }

        public virtual bool IsLabel(out ILabelDeletionTarget labelTarget)
        {
            labelTarget = null;
            return false;
        }

        public ParserRuleContext ScopingContext { get; }

        public override ParserRuleContext DeleteContext => AssociatedLabelToDelete != null
            ? TargetContext
            : TargetContext.GetChild<VBAParser.MainBlockStmtContext>();

        public virtual bool HasSameLogicalLineLabel(out VBAParser.StatementLabelDefinitionContext labelContext)
        {
            return TargetContext.TryGetChildContext(out labelContext);
        }

        public override string BuildEOSReplacementContent()
        {
            if (!(DeleteContext.Parent is ParserRuleContext prc
                && prc.TryGetChildContext<VBAParser.StatementLabelDefinitionContext>(out _)))
            {
                //No label to contend with
                return base.BuildEOSReplacementContent();
            }

            var replacement = string.Empty;
            var separationAndIndentation = string.Empty;

            if (!ModifiedTargetEOSContent.Contains(Tokens.CommentMarker))
            {
                var priorToSeparationContent = GetCurrentTextPriorToSeparationAndIndentation(PrecedingEOSContext, Rewriter);

                if (priorToSeparationContent.Contains(Tokens.CommentMarker))
                {
                    replacement = priorToSeparationContent;
                }

                separationAndIndentation = PrecedingEOSContext.GetSeparation();
            }

            return replacement + separationAndIndentation;
        }
    }
}
