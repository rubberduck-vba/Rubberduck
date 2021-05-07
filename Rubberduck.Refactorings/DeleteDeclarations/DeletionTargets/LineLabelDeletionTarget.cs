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
    internal class LineLabelDeletionTarget : DeleteDeclarationTarget, ILineLabelDeletionTarget
    {
        public LineLabelDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target)
            : base(declarationFinderProvider, target)
        {
            _listContext = null;

            _targetContext = target.Context.GetAncestor<VBAParser.BlockStmtContext>();

            //If there is a declaration on the label's line, delete just the label's context
            _deleteContext = _targetContext.TryGetChildContext<VBAParser.MainBlockStmtContext>(out _)
                ? target.Context
                : target.Context.GetAncestor<VBAParser.BlockStmtContext>();

            _eosContext = _targetContext.TryGetChildContext<VBAParser.MainBlockStmtContext>(out _)
                ? null
                : GetFollowingEndOfStatementContext(_targetContext);
        }
    }
}
