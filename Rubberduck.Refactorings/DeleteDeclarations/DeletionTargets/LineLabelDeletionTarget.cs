using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class LineLabelDeletionTarget : ProcedureLocalDeletionTarget<VBAParser.IdentifierStatementLabelContext>, ILabelDeletionTarget
    {
        public LineLabelDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
            : base(declarationFinderProvider, target, rewriter)
        {
            ListContext = null;

            DeleteContext = target.Context.GetAncestor<VBAParser.BlockStmtContext>();

            TargetEOSContext = TargetContext.GetFollowingEndOfStatementContext();

            if (TargetContext.TryGetChildContext<VBAParser.MainBlockStmtContext>(out _))
            {
                //There is a declaration on the same logical line - delete just the label
                //Note: ProcedureLocalDeletionTarget handles cases where a Label AND a declaration on the same logical line are to be deleted  
                DeleteContext = target.Context;
                TargetEOSContext = null;
            }
        }

        public override bool IsLabel(out ILabelDeletionTarget labelTarget)
        {
            labelTarget = this;
            return true;
        }

        public override ParserRuleContext DeleteContext { protected set; get; }

        public bool HasSameLogicalLineListContext(out ParserRuleContext varOrConst)
        {
            varOrConst = null;
            if (TargetContext.TryGetChildContext<VBAParser.MainBlockStmtContext>(out var relatedVarOrConst))
            {
                if (relatedVarOrConst.TryGetChildContext<VBAParser.ConstStmtContext>(out var constStmtContext))
                {
                    varOrConst = constStmtContext;
                }
                else if (relatedVarOrConst.TryGetChildContext<VBAParser.VariableStmtContext>(out var varStmtContext))
                {
                    varOrConst = varStmtContext.GetChild<VBAParser.VariableListStmtContext>();
                }
            }
            return varOrConst != null;
        }

        public override bool HasSameLogicalLineLabel(out VBAParser.StatementLabelDefinitionContext labelContext)
        {
            labelContext = null;
            return false;
        }

        public bool ReplaceLabelWithWhitespace { set; get; }

        public bool HasFollowingMainBlockStatementContext(out VBAParser.MainBlockStmtContext mainBlockStmtContext)
        {
            mainBlockStmtContext = TargetContext.GetDescendent<VBAParser.MainBlockStmtContext>();
            return mainBlockStmtContext != null;
        }

    }
}
