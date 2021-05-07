using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    public class DeleteModuleElementsRefactoringAction : DeleteElementRefactoringActionBase<DeleteModuleElementsModel>
    {
        public DeleteModuleElementsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
            : base(declarationFinderProvider, rewritingManager)
        {}

        public override void Refactor(DeleteModuleElementsModel model, IRewriteSession rewriteSession)
        {
            if (model.Targets.Any(t =>!( t.ParentDeclaration is ModuleDeclaration)))
            {
                throw new InvalidOperationException("Only module-level declarations can be refactored by this object");
            }

            var targetsByQMN = model.Targets.ToLookup(t => t.QualifiedModuleName);

            var qmnTargets = new Dictionary<QualifiedModuleName, List<IDeclarationDeletionTarget>>();


            foreach (var targetGroup in targetsByQMN)
            {
                qmnTargets.Add(targetGroup.Key, new List<IDeclarationDeletionTarget>());

                var targetsInSameModule = targetGroup.ToList();

                var members = targetsInSameModule.Where(d => d.DeclarationType.HasFlag(DeclarationType.Member));
                var targetsContainedInTargetMembers = targetsInSameModule.Where(t => members.Contains(t.ParentDeclaration));

                targetsInSameModule.RemoveAll(t => targetsContainedInTargetMembers.Contains(t));

                while (targetsInSameModule.Any())
                {
                    targetsInSameModule = CreateDeleteDeclarationTarget(targetsInSameModule, out var deleteTarget);
                    qmnTargets[targetGroup.Key].Add(deleteTarget);
                }
            }

            foreach (var qmn in qmnTargets.Keys)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(qmn);
                DeleteTargetsInModule(qmnTargets[qmn], rewriter);
            }
        }

        private List<Declaration> CreateDeleteDeclarationTarget(List<Declaration> targets, out IDeclarationDeletionTarget deleteTarget)
        {
            var remainingTargets = targets;

            var target = targets.First();

            deleteTarget = new ModuleElementDeletionTarget(_declarationFinderProvider, target);

            if (deleteTarget.AllDeclarationsInListContext.Count >= 1)
            {
                var listContextRelatedTargets = deleteTarget.AllDeclarationsInListContext.Intersect(targets);
                deleteTarget.AddTargets(listContextRelatedTargets);
                remainingTargets.RemoveAll(t => listContextRelatedTargets.Contains(t));
            }
            else
            {
                remainingTargets.RemoveAll(t => t == target);
            }


            return remainingTargets;
        }
        private void DeleteTargetsInModule(IEnumerable<IDeclarationDeletionTarget> allTargets, IModuleRewriter rewriter)
        {
            var moduleElementDeleteTargets = allTargets
                .Where(dt => !(dt is IProcedureLocalDeletionTarget
                    || dt is IEnumMemberDeletionTarget
                    || dt is IUdtMemberDeletionTarget
                    || dt is ILineLabelDeletionTarget)).ToList();

            if (!moduleElementDeleteTargets.Any())
            {
                return;
            }

            (ParserRuleContext declarationSection, ParserRuleContext codeSection) = GetModuleDeclarationAndCodeSectionContexts(moduleElementDeleteTargets.First().TargetProxy);

            var moduleDeclarationElements = declarationSection.children?
                .Where(ch => ch is VBAParser.ModuleDeclarationsElementContext)
                .Cast<ParserRuleContext>() ?? Enumerable.Empty<VBAParser.ModuleDeclarationsElementContext>();

            var moduleBodyElements = codeSection.children?
                .Where(ch => ch is VBAParser.ModuleBodyElementContext)
                .Cast<ParserRuleContext>() ?? Enumerable.Empty<VBAParser.ModuleBodyElementContext>();

            var orderedContexts = moduleDeclarationElements.Cast<ParserRuleContext>()
                .Concat(moduleBodyElements.Cast<ParserRuleContext>())
                .OrderBy(c => c.GetSelection());

            var deletionGroups = CreateDeletionGroups(moduleElementDeleteTargets, orderedContexts);

            foreach (var deletionGroup in deletionGroups)
            {
                DeleteGroup(deletionGroup, moduleElementDeleteTargets, rewriter);
                RemovePartialDeletions(deletionGroup, moduleElementDeleteTargets, rewriter);
            }
        }

        private void RemovePartialDeletions(DeletionGroup deletionGroup, List<IDeclarationDeletionTarget> deleteTargets, IModuleRewriter rewriter)
        {
            foreach (var de in deletionGroup.Contexts)
            {
                var decDeleteTarget = deleteTargets.FirstOrDefault(d => d.TargetContext == de);
                if (decDeleteTarget?.IsFullDelete ?? true)
                {
                    continue;
                }

                RemoveListDeclarationSubsetVariableOrConstant(decDeleteTarget, rewriter);
            }
        }

        protected override void ModifyEOSContexts(DeletionGroup deletionGroup, IDeclarationDeletionTarget decDeleteTarget, IModuleRewriter rewriter)
        {
            if (deletionGroup.PrecedingNonDeletedContext != null)
            {
                deletionGroup.PrecedingNonDeletedContext.TryGetFollowingContext<VBAParser.EndOfStatementContext>(out var eos);
                (decDeleteTarget as IModuleElementDeletionTarget).SetPrecedingEOSContext(eos);
            }

            base.ModifyEOSContexts(deletionGroup, decDeleteTarget, rewriter);
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
            var moduleBodyCtxt = (moduleDeclarationsContext.Parent as ParserRuleContext).GetChild<VBAParser.ModuleBodyContext>();
            return (moduleDeclarationsContext, moduleBodyCtxt);
        }
    }
}
