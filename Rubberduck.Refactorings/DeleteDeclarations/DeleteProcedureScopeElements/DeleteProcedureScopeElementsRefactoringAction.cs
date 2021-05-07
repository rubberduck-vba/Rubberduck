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
    public class DeleteProcedureScopeElementsRefactoringAction : DeleteElementRefactoringActionBase<DeleteProcedureScopeElementsModel>
    {
        public DeleteProcedureScopeElementsRefactoringAction(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager)
           : base(declarationFinderProvider, rewritingManager)
        {

        }
        public override void Refactor(DeleteProcedureScopeElementsModel model, IRewriteSession rewriteSession)
        {
            if (model.Targets.Any(t => t.ParentDeclaration is ModuleDeclaration))
            {
                throw new InvalidOperationException("Only declarations within procedures can be refactored by this object");
            }

            var targetsByQMN = model.Targets.ToLookup(t => t.QualifiedModuleName);

            var qmnTargets = new Dictionary<QualifiedModuleName, List<IDeclarationDeletionTarget>>();

            foreach (var targetGroup in targetsByQMN)
            {
                qmnTargets.Add(targetGroup.Key, new List<IDeclarationDeletionTarget>());

                var targetsInSameModule = targetGroup.ToList();

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

        public void DeleteTargetsInModule(IEnumerable<IDeclarationDeletionTarget> allTargets, IModuleRewriter rewriter)
        {
            var blockDeleteTargets = allTargets
                .Where(dt => dt is IProcedureLocalDeletionTarget
                    || dt is ILineLabelDeletionTarget).ToList();

            if (!blockDeleteTargets.Any())
            {
                return;
            }

            blockDeleteTargets = HandleLabelAndVarOrConstInSameBlock(blockDeleteTargets);

            var contexts = GetAllContextElements<VBAParser.BlockContext, VBAParser.BlockStmtContext>(blockDeleteTargets.First().TargetProxy);

            var deletionGroups = CreateDeletionGroups(blockDeleteTargets, contexts);

            foreach (var deletionGroup in deletionGroups)
            {
                DeleteGroup(deletionGroup, blockDeleteTargets, rewriter);
                RemovePartialDeletions(deletionGroup, blockDeleteTargets, rewriter);
            }
        }

        private List<Declaration> CreateDeleteDeclarationTarget(List<Declaration> targets, out IDeclarationDeletionTarget deleteTarget)
        {
            var remainingTargets = targets;

            var target = targets.First();


            switch (target.DeclarationType)
            {
                case DeclarationType.Variable:
                    deleteTarget = new ProcedureLocalDeletionTarget<VBAParser.VariableListStmtContext>(_declarationFinderProvider, target);
                    break;
                case DeclarationType.Constant:
                    deleteTarget = new ProcedureLocalDeletionTarget<VBAParser.ConstStmtContext>(_declarationFinderProvider, target);
                    break;
                case DeclarationType.LineLabel:
                    deleteTarget = new LineLabelDeletionTarget(_declarationFinderProvider, target);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported DeclarationType: {target.DeclarationType}");
            }

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

        private static List<IDeclarationDeletionTarget> HandleLabelAndVarOrConstInSameBlock(List<IDeclarationDeletionTarget> blockDeleteTargets)
        {
            var check = blockDeleteTargets.ToLookup(k => k.TargetContext);
            if (check.Count < blockDeleteTargets.Count)
            {
                var newList = new List<IDeclarationDeletionTarget>();
                foreach (var group in check)
                {
                    var varOrConst = group.Single(e => e is IProcedureLocalDeletionTarget) as IProcedureLocalDeletionTarget;
                    newList.Add(varOrConst);
                    foreach (var element in group)
                    {
                        if (element is ILineLabelDeletionTarget labelTarget)
                        {
                            varOrConst.DeleteAssociatedLabel = true;
                        }
                    }
                    return newList;
                }
            }
            return blockDeleteTargets;
        }

        protected override void ModifyEOSContextForLastTargetOfDeletionGroup(IDeclarationDeletionTarget deleteTarget, IModuleRewriter rewriter)
        {
            //Modify the proceding non-deleted EOS with EOS content from the last deleted target.

            var groupEndingEOSContentProvider = new EOSContextContentProvider(deleteTarget.EndOfStatementContext, rewriter);

            if (groupEndingEOSContentProvider.ModifiedEOSContent.StartsWith(": "))
            {
                rewriter.Remove(deleteTarget.EndOfStatementContext);
                return;
            }

            var precedingEOSContentProvider = new EOSContextContentProvider(deleteTarget.PrecedingEOSContext, rewriter);

            string replacementContent;

            if ((deleteTarget.DeleteContext.Parent as ParserRuleContext).TryGetChildContext<VBAParser.StatementLabelDefinitionContext>(out var labelContext))
            {
                if (groupEndingEOSContentProvider.ModifiedContentContainsCommentMarker)
                {
                    replacementContent = precedingEOSContentProvider.ModifiedContentContainsCommentMarker
                        ? $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{labelContext.GetText()}{groupEndingEOSContentProvider.ModifiedEOSContent}"
                        : $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{precedingEOSContentProvider.Separation}{labelContext.GetText()}{groupEndingEOSContentProvider.ModifiedEOSContent}";
                }
                else
                {
                    //Deletes only the variable and not the label.  Appends the EOS content to the remaining
                    //content of the declaration line (the label) and replaces the target context - the prior EOS 
                    //is injected as part of the TargetContext replacement
                    var modifedContent = GetModifiedContextText(deleteTarget.DeleteContext.Parent as ParserRuleContext, rewriter);
                    replacementContent = $"{modifedContent}{groupEndingEOSContentProvider.Separation}{groupEndingEOSContentProvider.Indentation}";

                    rewriter.Replace(deleteTarget.TargetContext, replacementContent);
                    rewriter.Remove(deleteTarget.EndOfStatementContext);
                    rewriter.Replace(deleteTarget.PrecedingEOSContext, $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{precedingEOSContentProvider.Separation}");
                    return;
                }
            }
            else
            {
                if (groupEndingEOSContentProvider.ModifiedContentContainsCommentMarker)
                {
                    replacementContent = precedingEOSContentProvider.ModifiedContentContainsCommentMarker
                        ? $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation}{groupEndingEOSContentProvider.ModifiedEOSContent}"//.Substring(Environment.NewLine.Length)}"
                        : $"{precedingEOSContentProvider.ContentPriorToSeparationAndIndentation.Trim()}{precedingEOSContentProvider.Separation}{groupEndingEOSContentProvider.ContentFreeOfStartingNewLines}";
                }
                else
                {
                    var replacementFirstPart = precedingEOSContentProvider.ContentPriorToSeparationAndIndentation.Contains("'")
                        ? precedingEOSContentProvider.ContentPriorToSeparationAndIndentation
                        : precedingEOSContentProvider.ContentPriorToSeparationAndIndentation.Trim();

                    replacementContent = $"{replacementFirstPart}{precedingEOSContentProvider.Separation}{groupEndingEOSContentProvider.Indentation}";
                }
            }

            //Replace the precedingEOSContent if it exists, or modify the last deleted target's EOSContext
            if (deleteTarget.PrecedingEOSContext != null)
            {
                rewriter.Replace(deleteTarget.PrecedingEOSContext, replacementContent);
                if (deleteTarget.EndOfStatementContext != null)
                {
                    if (labelContext != null)
                    {
                        rewriter.Replace(deleteTarget.EndOfStatementContext, groupEndingEOSContentProvider.OriginalEOSContent);
                    }
                    else
                    {
                        rewriter.Remove(deleteTarget.EndOfStatementContext);
                    }
                }
            }
            else if (deleteTarget.EndOfStatementContext != null)
            {
                rewriter.Replace(deleteTarget.EndOfStatementContext, replacementContent);
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

                //TODO: Need test(s) for Label with same-line multi variable list that is fully and partially deleted
                RemoveListDeclarationSubsetVariableOrConstant(decDeleteTarget, rewriter);
            }
        }
    }
}
