using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.DeleteDeclarations.Abstract;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class EnumMemberDeletionTarget : DeclarationDeletionTargetBase, IEnumMemberDeletionTarget
    {
        public EnumMemberDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
            : base(declarationFinderProvider, target, rewriter)
        {
            ListContext = target.Context.GetAncestor<VBAParser.EnumerationStmtContext>();

            TargetContext = target.Context;

            DeleteContext = target.Context.GetChild<VBAParser.IdentifierContext>();

            PrecedingEOSContext = TargetProxy.Context == ListContext.children.SkipWhile(ch => !(ch is VBAParser.EnumerationStmt_ConstantContext)).First()
                ? ListContext.GetChild<VBAParser.EndOfStatementContext>()
                : (ListContext.children
                    .TakeWhile(ch => ch != TargetProxy.Context)
                    .Last() as ParserRuleContext)
                    .GetChild<VBAParser.EndOfStatementContext>();

            TargetEOSContext = DeleteContext.GetFollowingEndOfStatementContext();
        }
    }
}
