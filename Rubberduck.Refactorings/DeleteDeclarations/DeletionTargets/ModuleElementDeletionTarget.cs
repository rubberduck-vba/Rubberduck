using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class ModuleElementDeletionTarget : DeclarationDeletionTargetBase, IModuleElementDeletionTarget
    {
        public ModuleElementDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
            : base(declarationFinderProvider, target, rewriter)
        {
            ListContext = GetListContext(target);

            if (target.Context.TryGetAncestor<VBAParser.ModuleDeclarationsElementContext>(out var mde))
            {
                TargetContext = mde;
            }
            else if (target.Context.TryGetAncestor<VBAParser.ModuleBodyElementContext>(out var mbe))
            {
                TargetContext = mbe;
            }

            DeleteContext = target.DeclarationType.HasFlag(DeclarationType.Member)
                ? target.Context.GetAncestor<VBAParser.ModuleBodyElementContext>()
                : target.Context.GetAncestor<VBAParser.ModuleDeclarationsElementContext>() as ParserRuleContext;

            //The preceding EOS Context cannot be determined directly from the target.  It depends upon what else is deleted
            //adjacent to the target.
            PrecedingEOSContext = null;

            TargetEOSContext = DeleteContext.GetFollowingEndOfStatementContext();
        }

        public override bool IsFullDelete
            => TargetProxy.DeclarationType != DeclarationType.Variable && TargetProxy.DeclarationType != DeclarationType.Constant
                || AllDeclarationsInListContext.Intersect(Targets).Count() == AllDeclarationsInListContext.Count;
    }
}
