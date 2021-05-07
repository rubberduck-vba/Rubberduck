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
    internal class EnumMemberDeletionTarget : DeleteDeclarationTarget, IEnumMemberDeletionTarget
    {
        public EnumMemberDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target)
            : base(declarationFinderProvider, target)
        {
            _targetContext = target.Context;

            _deleteContext = target.Context.GetChild<VBAParser.IdentifierContext>();

            _listContext = target.Context.GetAncestor<VBAParser.EnumerationStmtContext>();

            var enumStmtContext = TargetProxy.Context.GetAncestor<VBAParser.EnumerationStmtContext>();
            _precedingEOSContext = TargetProxy.Context == enumStmtContext.children.SkipWhile(ch => !(ch is VBAParser.EnumerationStmt_ConstantContext)).First()
                ? enumStmtContext.GetChild<VBAParser.EndOfStatementContext>()
                : (enumStmtContext.children
                    .TakeWhile(ch => ch != TargetProxy.Context)
                    .Last() as ParserRuleContext)
                    .GetChild<VBAParser.EndOfStatementContext>();

            _eosContext = GetFollowingEndOfStatementContext(_deleteContext);
        }
    }
}
