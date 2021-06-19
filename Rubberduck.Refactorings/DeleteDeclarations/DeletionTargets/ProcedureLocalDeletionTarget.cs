using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class ProcedureLocalDeletionTarget<T> : DeleteDeclarationTarget, IProcedureLocalDeletionTarget where T : ParserRuleContext
    {
        public ProcedureLocalDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target)
            : base(declarationFinderProvider, target)
        {
            _listContext = target.Context.GetAncestor<T>();

            //If there is a label on the declaration's line, delete just the declaration's context
            _targetContext = target.Context.GetAncestor<VBAParser.BlockStmtContext>();

            _deleteContext = _targetContext.TryGetChildContext<VBAParser.StatementLabelDefinitionContext>(out _)
                ? _targetContext.GetChild<VBAParser.MainBlockStmtContext>()
                : target.Context.GetAncestor<VBAParser.BlockStmtContext>() as ParserRuleContext;

            _eosContext = GetFollowingEndOfStatementContext(_deleteContext);

            _precedingEOSContext = GetEndOfStatementContext(TargetProxy);
        }

        public override VBAParser.EndOfStatementContext PrecedingEOSContext => _precedingEOSContext;

        public override bool IsFullDelete 
            => _allDeclarationsInList.Intersect(_targets).Count() == _allDeclarationsInList.Count;
        
        public bool DeleteAssociatedLabel { set; get; } = false;

        public override ParserRuleContext DeleteContext => DeleteAssociatedLabel 
            ? _targetContext 
            : _targetContext.GetChild<VBAParser.MainBlockStmtContext>();

        public override bool HasPrecedingLabel(out VBAParser.StatementLabelDefinitionContext labelContext)
        {
            labelContext = null;
            var result =  DeleteContext.Parent is ParserRuleContext prc && prc.TryGetChildContext(out labelContext);
            return result;
        }


        private static VBAParser.EndOfStatementContext GetEndOfStatementContext(Declaration target)
        {
            var blockContext = target.Context.GetAncestor<VBAParser.BlockContext>();
            var targetBlockStmt = target.Context.GetAncestor<VBAParser.BlockStmtContext>();
            VBAParser.EndOfStatementContext precedingEOSContext;

            precedingEOSContext = blockContext.children
                .TakeWhile(ch => !(ch is VBAParser.BlockStmtContext bst && bst == targetBlockStmt))
                .LastOrDefault() as VBAParser.EndOfStatementContext;

            //precedingEOSContext will be null if the target is the first Declaration following the procedure Declaration
            if (precedingEOSContext is null)
            {
                var arglistCtxt = target.Context
                    .GetAncestor<VBAParser.ModuleBodyElementContext>()
                    .GetDescendent<VBAParser.ArgListContext>();
                precedingEOSContext = GetFollowingEndOfStatementContext(arglistCtxt);
            }
            return precedingEOSContext;
        }
    }
}
