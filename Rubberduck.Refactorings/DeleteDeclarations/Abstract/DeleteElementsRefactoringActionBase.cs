using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Rubberduck.Refactorings.DeleteDeclarations.Abstract
{
    public abstract class DeleteElementsRefactoringActionBase<TModel> : CodeOnlyRefactoringActionBase<TModel> where TModel : class, IRefactoringModel
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IDeclarationDeletionTargetFactory _declarationDeletionTargetFactory;
        private readonly IDeclarationDeletionGroupsGenerator _declarationDeletionGroupsGenerator;
        
        private static readonly string _lineContinuationExpression = $"{Tokens.LineContinuation}{Environment.NewLine}";

        protected const string EOS_COLON = ": ";

        public DeleteElementsRefactoringActionBase(IDeclarationFinderProvider declarationFinderProvider,
            IDeclarationDeletionTargetFactory deletionTargetFactory,
            IDeclarationDeletionGroupsGeneratorFactory deletionGroupsGeneratorFactory,
            IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _declarationDeletionTargetFactory = deletionTargetFactory;
            _declarationDeletionGroupsGenerator = deletionGroupsGeneratorFactory.Create();
        }

        protected abstract bool CanRefactorAllTargets(TModel model);

        protected void DeleteDeclarations(IDeleteDeclarationsModel model, 
            IRewriteSession rewriteSession, 
            Func<IEnumerable<Declaration>, IRewriteSession, IDeclarationDeletionTargetFactory, IEnumerable<IDeclarationDeletionTarget>> generateDeletionTargets)
        {
            var deletionTargets = generateDeletionTargets(model.Targets, rewriteSession, _declarationDeletionTargetFactory);

            var targetsLookup = deletionTargets.ToLookup(dt => dt.TargetProxy.QualifiedModuleName);

            foreach (var moduleQualifiedDeleteGroups in targetsLookup)
            {
                var deletionGroups = _declarationDeletionGroupsGenerator.Generate(moduleQualifiedDeleteGroups);

                var rewriter = rewriteSession.CheckOutModuleRewriter(moduleQualifiedDeleteGroups.Key);

                if (!model.DeleteDeclarationsOnly && model.DeleteAnnotations)
                {
                    DeleteAnnotations(_declarationFinderProvider, deletionGroups, rewriter);
                }

                RemoveDeletionGroups(deletionGroups, model, rewriter);
            }
        }

        protected void RemoveDeletionGroups(IEnumerable<IDeclarationDeletionGroup> deletionGroups, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            foreach (var deletionGroup in deletionGroups)
            {
                if (deletionGroup.OrderedPartialDeletionTargets.Any())
                {
                    RemovePartialDeletionTargets(deletionGroup, model, rewriter);
                }

                if (deletionGroup.OrderedFullDeletionTargets.Any())
                {
                    RemoveFullDeletionGroup(deletionGroup, model, rewriter);
                }
            }
        }

        protected void RemoveFullDeletionGroup(IDeclarationDeletionGroup deletionGroup, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            foreach (var deleteTarget in deletionGroup.OrderedFullDeletionTargets)
            {
                DeleteTarget(deleteTarget, rewriter);
            }

            if (model.DeleteDeclarationsOnly)
            {
                return;
            }

            var lastTarget = deletionGroup.OrderedFullDeletionTargets.LastOrDefault();

            foreach (var deleteTarget in deletionGroup.OrderedFullDeletionTargets.Where(t => t != lastTarget && t.TargetEOSContext != null))
            {
                rewriter.Remove(deleteTarget.TargetEOSContext);
            }

            if (lastTarget is null || lastTarget.TargetEOSContext is null)
            {
                return;
            }

            lastTarget.PrecedingEOSContext = GetPrecedingNonDeletedEOSContextForGroup(deletionGroup);

            ModifyLastTargetEOS(lastTarget, model, rewriter);
        }

        // The default GetPrecedingNonDeletedEOSContextForGroup is overridden by DeleteModuleElementsRefactoringAction 
        // and DeleteProcedureScopeElementsRefactoringAction
        protected virtual VBAParser.EndOfStatementContext GetPrecedingNonDeletedEOSContextForGroup(IDeclarationDeletionGroup deletionGroup)
            => deletionGroup.Targets.FirstOrDefault()?.PrecedingEOSContext;

        protected IEnumerable<IDeclarationDeletionTarget> CreateDeletionTargetsSupportingPartialDeletions(IEnumerable<Declaration> declarations, IRewriteSession rewriteSession, IDeclarationDeletionTargetFactory targetFactory)
        {
            var deletionTargets = new List<IDeclarationDeletionTarget>();

            var remainingTargets = declarations.ToList();

            while (remainingTargets.Any())
            {
                var deleteTarget = targetFactory.Create(remainingTargets.First(), rewriteSession);

                if (deleteTarget.AllDeclarationsInListContext.Count >= 1)
                {
                    var listContextRelatedTargets = deleteTarget.AllDeclarationsInListContext.Intersect(declarations);
                    deleteTarget.AddTargets(listContextRelatedTargets);
                    remainingTargets.RemoveAll(t => listContextRelatedTargets.Contains(t));
                }
                else
                {
                    remainingTargets.RemoveAll(t => t == declarations.First());
                }

                deletionTargets.Add(deleteTarget);
            }

            return deletionTargets;
        }

        protected Action<IDeclarationDeletionTarget, IModuleRewriter> DeleteTarget { set; get; }
            = (t, rewriter) => rewriter.Remove(t.DeleteContext);

        /// <summary>
        /// Replaces the EndOfStatementContext preceding the deletion group. 
        /// </summary>
        /// <remarks>
        /// The preceding EndOfStatementContext is replaced with a merged version of the preceding EndOfStatementContext
        /// and the last delete target's EndOfStatementContext.  
        /// </remarks>
        protected void ModifyLastTargetEOS(IDeclarationDeletionTarget lastTarget, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            if (lastTarget.TargetEOSContext.GetText() == EOS_COLON)
            {
                //Remove the declarations EOS colon character and use the PrecedingEOSContext as-is
                lastTarget.Rewriter.Remove(lastTarget.TargetEOSContext);
                return;
            }

            ModifyRelatedComments(lastTarget, model, rewriter);

            var replacementText = lastTarget.EOSContextToReplace == lastTarget.TargetEOSContext
                ? lastTarget.ModifiedTargetEOSContent
                : lastTarget.BuildEOSReplacementContent();

            rewriter.Replace(lastTarget.EOSContextToReplace, replacementText);

            if (lastTarget.DeletionIncludesEOSContext)
            {
                rewriter.Remove(lastTarget.TargetEOSContext);
            }
        }

        protected static void ModifyRelatedComments(IDeclarationDeletionTarget deleteTarget, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            var targetEOSComments = deleteTarget.TargetEOSContext.GetAllComments();

            if (deleteTarget.IsFullDelete)
            {
                var declarationLogicalLineCommentContext = deleteTarget.GetDeclarationLogicalLineCommentContext();

                if (model.DeleteDeclarationLogicalLineComments && declarationLogicalLineCommentContext != null)
                {
                    DeleteDeclarationLogicalLineComments(deleteTarget, declarationLogicalLineCommentContext, rewriter);
                    targetEOSComments = targetEOSComments.Where(c => c != declarationLogicalLineCommentContext);
                }
                else if (!model.DeleteDeclarationLogicalLineComments && declarationLogicalLineCommentContext != null)
                {
                    //If we are keeping the Declaration line comments, then insert a newline or it will end up on the
                    //same line as the last comment of the preceding EOSContext
                    rewriter.InsertBefore(declarationLogicalLineCommentContext.Start.TokenIndex, Environment.NewLine);
                }
            }

            if (model.InsertValidationTODOForRetainedComments)
            {
                var injectedTODOContent = Resources.Refactorings.Refactorings.CommentVerification_TODO;

                foreach (var comment in targetEOSComments.Concat(deleteTarget.PrecedingEOSContext.GetAllComments()))
                {
                    var content = comment.GetText();
                    var indexOfFirstCommentMarker = content.IndexOf(Tokens.CommentMarker);
                    var newContent = $"{content.Substring(0, indexOfFirstCommentMarker)}{injectedTODOContent}{content.Substring(indexOfFirstCommentMarker + 1)}";
                    rewriter.Replace(comment, newContent);
                }
            }
        }

        /// <summary>
        /// Deletes only those Annotations where ALL the Declarations referencing the same Annotation are selected for deletion.  
        /// </summary>
        private static void DeleteAnnotations(IDeclarationFinderProvider declarationFinderProvider, IReadOnlyCollection<IDeclarationDeletionGroup> deletionGroups, IModuleRewriter rewriter)
        {
            foreach (var deletionGroup in deletionGroups)
            {
                if (!TryGetDeletableAnnotations(deletionGroup, declarationFinderProvider, out var deletableAnnotations))
                {
                    continue;
                }

                foreach (var annotation in deletableAnnotations)
                {
                    if (annotation.TryGetAncestor<VBAParser.IndividualNonEOFEndOfStatementContext>(out var annotationListIndividualNonEOFEOSCtxt))
                    {
                        rewriter.Remove(annotationListIndividualNonEOFEOSCtxt);
                    }
                }
            }
        }
        private static bool TryGetDeletableAnnotations(IDeclarationDeletionGroup deletionGroup, IDeclarationFinderProvider declarationFinderProvider, out List<VBAParser.AnnotationContext> deletableAnnotations)
        {
            deletableAnnotations = new List<VBAParser.AnnotationContext>();

            var relevantAnnotations = deletionGroup.Declarations
                .SelectMany(d => d.Annotations)
                .Select(a => a.Context)
                .Distinct();

            var moduleDeclarations = declarationFinderProvider.DeclarationFinder
                .Members(deletionGroup.Declarations.First().QualifiedModuleName).ToList();

            foreach (var annotation in relevantAnnotations)
            {
                var declarationsAssociatedWithAnnotation = moduleDeclarations
                    .Where(t => t.Annotations.Any(a => a.Context == annotation));

                if (declarationsAssociatedWithAnnotation.Any(d => !deletionGroup.Declarations.Contains(d)))
                {
                    continue;
                }

                deletableAnnotations.Add(annotation);
            }

            return deletableAnnotations.Any();
        }

        private static void DeleteDeclarationLogicalLineComments(IDeclarationDeletionTarget deleteTarget, VBAParser.CommentContext declarationLineCommentContext, IModuleRewriter rewriter)
        {
            if (declarationLineCommentContext is null)
            {
                return;
            }

            var individualNonEOFEOS = declarationLineCommentContext.GetAncestor<VBAParser.IndividualNonEOFEndOfStatementContext>();
            var contextToDelete = individualNonEOFEOS.GetChild<VBAParser.EndOfLineContext>();
            
            var ws = contextToDelete.GetDescendent<VBAParser.WhiteSpaceContext>();
            var containsLineContinuation = ws?.GetText().Contains(Tokens.LineContinuation) ?? false;

            if (contextToDelete != null && declarationLineCommentContext.Start.Line == deleteTarget.TargetEOSContext.Start.Line || containsLineContinuation)
            {
                rewriter.Remove(contextToDelete);
            }
        }

        private void RemovePartialDeletionTargets(IDeclarationDeletionGroup deletionGroup, IDeleteDeclarationsModel model, IModuleRewriter rewriter)
        {
            var lastTarget = deletionGroup.OrderedPartialDeletionTargets.Last();

            lastTarget.PrecedingEOSContext = GetPrecedingNonDeletedEOSContextForGroup(deletionGroup);

            var retainedDeclarationsExpression = lastTarget.ListContext.GetText().Contains(_lineContinuationExpression)
                ? $"{BuildDeclarationsExpressionWithLineContinuations(lastTarget)}"
                : $"{string.Join(", ", lastTarget.RetainedDeclarations.Select(d => d.Context.GetText()))}";

            rewriter.Replace(lastTarget.ListContext.Parent, $"{GetDeclarationScopeExpression(lastTarget.TargetProxy)} {retainedDeclarationsExpression}");

            if (lastTarget is null || lastTarget.TargetEOSContext is null)
            {
                return;
            }

            if (lastTarget.TargetEOSContext.GetText() == EOS_COLON)
            {
                //Remove the declarations EOS colon character and use the PrecedingEOSContext as-is
                lastTarget.Rewriter.Remove(lastTarget.TargetEOSContext);
                return;
            }

            ModifyRelatedComments(lastTarget, model, rewriter);

            rewriter.Replace(lastTarget.TargetEOSContext, lastTarget.ModifiedTargetEOSContent);
        }

        private static string BuildDeclarationsExpressionWithLineContinuations(IDeclarationDeletionTarget deleteDeclarationTarget)
        {
            var elementsByLineContinuation = deleteDeclarationTarget.ListContext.GetText().Split(new string[] { _lineContinuationExpression }, StringSplitOptions.None);

            if (elementsByLineContinuation.Count() == 1)
            {
                throw new ArgumentException("'targetsToDelete' parameter does not contain line extension(s)");
            }

            var expr = new StringBuilder();
            foreach (var element in elementsByLineContinuation)
            {
                var idContexts = deleteDeclarationTarget.RetainedDeclarations.Where(r => element.Contains(r.Context.GetText())).Select(d => d);
                foreach (var ctxt in idContexts)
                {
                    var indent = string.Concat(element.TakeWhile(e => e == ' '));

                    expr = expr.Length == 0
                        ? expr.Append(ctxt.Context.GetText())
                        : expr.Append($",{_lineContinuationExpression}{indent}{ctxt.Context.GetText()}");
                }
            }
            return expr.ToString();
        }

        private static string GetDeclarationScopeExpression(Declaration listPrototype)
        {
            if (listPrototype.DeclarationType.HasFlag(DeclarationType.Variable))
            {
                var accessToken = listPrototype.Accessibility == Accessibility.Implicit
                    ? Tokens.Private
                    : $"{listPrototype.Accessibility}";

                return listPrototype.ParentDeclaration is ModuleDeclaration
                    ? accessToken
                    : Tokens.Dim;
            }

            if (listPrototype.DeclarationType.HasFlag(DeclarationType.Constant))
            {
                var accessToken = listPrototype.Accessibility == Accessibility.Implicit
                    ? Tokens.Private
                    : $"{listPrototype.Accessibility}";

                return listPrototype.ParentDeclaration is ModuleDeclaration
                    ? $"{accessToken} {Tokens.Const}"
                    : Tokens.Const;
            }

            throw new ArgumentException("Unsupported DeclarationType");
        }
    }
}
