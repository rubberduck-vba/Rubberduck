using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteModuleElementsRefactoringAction : DeleteVariableOrConstantRefactoringActionBase<DeleteModuleElementsModel>
    {
        public DeleteModuleElementsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IDeclarationDeletionTargetFactory targetFactory, IDeleteDeclarationEndOfStatementContentModifierFactory eosModifierFactory, IRewritingManager rewritingManager)
            : base(declarationFinderProvider, targetFactory, eosModifierFactory, rewritingManager)
        {
            InjectRetrieveNonDeleteDeclarationForDeletionGroupAction(SetPrecedingNonDeletedEOSContextOfGroup);
        }

        public override void Refactor(DeleteModuleElementsModel model, IRewriteSession rewriteSession)
        {
            DeleteDeclarations(model, rewriteSession);
        }

        protected override IOrderedEnumerable<ParserRuleContext> GetAllContextElements(Declaration declaration)
        {
            (ParserRuleContext declarationSection, ParserRuleContext codeSection) = GetModuleDeclarationAndCodeSectionContexts(declaration);
            
            var moduleDeclarationElements = declarationSection?.children?
                .OfType<VBAParser.ModuleDeclarationsElementContext>()
                .Cast<ParserRuleContext>() ?? Enumerable.Empty<ParserRuleContext>();

            var moduleBodyElements = codeSection?.children?
                .OfType<VBAParser.ModuleBodyElementContext>()
                .Cast<ParserRuleContext>() ?? Enumerable.Empty<ParserRuleContext>();

            return moduleDeclarationElements
                .Concat(moduleBodyElements)
                .OrderBy(c => c.GetSelection());
        }

        private static void SetPrecedingNonDeletedEOSContextOfGroup(DeletionGroup deletionGroup, IDeclarationDeletionTarget deleteTarget)
        {
            //When building deletion groups for the Module Code Section, the result may point to the wrong 
            //preceding context if the first deleted Module Code Section element (of a DeletionGroup) is the first Member declaration of the module.
            
            //TODO: Add test for non-Option explicit with first Member deleted
            if (deleteTarget is IModuleElementDeletionTarget medt)
            {
                VBAParser.EndOfStatementContext eos = null;
                if (deletionGroup.PrecedingNonDeletedContext?.TryGetFollowingContext(out eos) ?? false)
                {
                    medt.SetPrecedingEOSContext(eos);
                }
            }
        }

        protected override void RefactorGuardClause(IDeleteDeclarationsModel model)
        {
            if (model.Targets.Any(t => !(t.ParentDeclaration is ModuleDeclaration)))
            {
                throw new InvalidOperationException("Only module-level declarations can be refactored by this object");
            }
        }

        private static (ParserRuleContext declarationSection, ParserRuleContext codeSection) GetModuleDeclarationAndCodeSectionContexts(Declaration target)
        {
            var bodyCtxt = target.Context.GetAncestor<VBAParser.ModuleBodyContext>();
            if (bodyCtxt != null)
            {
                var declarationsCtxt = (bodyCtxt.Parent as ParserRuleContext).GetChild<VBAParser.ModuleDeclarationsContext>();
                return (declarationsCtxt, bodyCtxt);
            }

            var moduleDeclarationsContext = target.Context.GetAncestor<VBAParser.ModuleDeclarationsContext>();
            var moduleBodyCtxt = (moduleDeclarationsContext?.Parent as ParserRuleContext).GetChild<VBAParser.ModuleBodyContext>();
            return (moduleDeclarationsContext, moduleBodyCtxt);
        }
    }
}
