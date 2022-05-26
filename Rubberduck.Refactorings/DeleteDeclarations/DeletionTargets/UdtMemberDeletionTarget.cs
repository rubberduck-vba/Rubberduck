using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class UdtMemberDeletionTarget : DeclarationDeletionTargetBase, IUdtMemberDeletionTarget
    {
        public UdtMemberDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
            : base(declarationFinderProvider, target, rewriter) 
        {
            ListContext = target.Context.GetAncestor<VBAParser.UdtMemberListContext>();

            TargetContext = target.Context;

            DeleteContext = target.Context;

            PrecedingEOSContext = DeleteContext == ListContext.children.First()
                ? TargetProxy.Context.GetAncestor<VBAParser.UdtDeclarationContext>()
                    .GetChild<VBAParser.EndOfStatementContext>()
                : ListContext.children.TakeWhile(ch => ch != TargetProxy.Context).Last() as VBAParser.EndOfStatementContext;

            TargetEOSContext = DeleteContext.GetFollowingEndOfStatementContext();
        }
    }
}

