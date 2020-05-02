using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberRefactoringSourcePreviewer : MoveMemberEndpointsPreviewer
    {
        public MoveMemberRefactoringSourcePreviewer(
            IRewritingManager rewritingManager,
            IMovedContentProviderFactory movedContentProviderFactory)
            : base(rewritingManager, movedContentProviderFactory) { }

        public override string PreviewMove(MoveMemberModel model)
            => PreviewEndpoints(model).source;
    }

    public class MoveMemberRefactoringDestinationPreviewer : MoveMemberEndpointsPreviewer
    {
        public MoveMemberRefactoringDestinationPreviewer(
            IRewritingManager rewritingManager,
            IMovedContentProviderFactory movedContentProviderFactory)
            : base(rewritingManager, movedContentProviderFactory) { }

        public override string PreviewMove(MoveMemberModel model)
            => PreviewEndpoints(model).destination;
    }

    public abstract class MoveMemberEndpointsPreviewer : IMoveMemberRefactoringPreviewer
    {
        private readonly IMovedContentProviderFactory _movedContentProviderFactory;
        private readonly IRewritingManager _rewritingManager;
        public MoveMemberEndpointsPreviewer(
            IRewritingManager rewritingManager,
            IMovedContentProviderFactory movedContentProviderFactory)
        {
            _movedContentProviderFactory = movedContentProviderFactory;
            _rewritingManager = rewritingManager;
        }

        public abstract string PreviewMove(MoveMemberModel model);

        protected (string source, string destination) PreviewEndpoints(MoveMemberModel model)
        {
            var destinationContent = string.Empty;

            var session = _rewritingManager.CheckOutCodePaneSession();
            var sourceRewriter = session.CheckOutModuleRewriter(model.Source.QualifiedModuleName);

            var sourceContent = LimitNewLines(sourceRewriter.GetText());

            var contentProvider = _movedContentProviderFactory.CreatePreviewProvider() as INewContentPreviewProvider;

            if (!model.TryGetStrategy(out var strategy))
            {
                destinationContent = contentProvider.AddNewContentDemarcation(Resources.RubberduckUI.MoveMember_ApplicableStrategyNotFound);
                return (sourceContent, destinationContent);
            }

            if (!model.SelectedDeclarations.Any())
            {
                destinationContent = contentProvider.AddNewContentDemarcation(Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove);
                return (sourceContent, destinationContent);
            }

            var refactorSession = _rewritingManager.CheckOutCodePaneSession();
            strategy.RefactorRewrite(model, refactorSession, _rewritingManager, contentProvider, out var newContent);

            var rewriter = refactorSession.CheckOutModuleRewriter(model.Source.QualifiedModuleName);
            sourceContent = LimitNewLines(rewriter.GetText());

            if (model.Destination.IsExistingModule(out var destinationModule))
            {
                rewriter = refactorSession.CheckOutModuleRewriter(destinationModule.QualifiedModuleName);
                destinationContent = LimitNewLines(rewriter.GetText());
            }
            else
            {
                var optionExplicit = $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}";
                destinationContent = $"{optionExplicit}{Environment.NewLine}{LimitNewLines(newContent)}";
            }

            return (sourceContent, destinationContent);
        }

        private static string LimitNewLines(string text, int maxConsecutiveNewLines = 2)
        {
            var target = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines + 1).ToList());
            var replacement = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            for (var counter = 1; counter < 50 && text.Contains(target); counter++)
            {
                text = text.Replace(target, replacement);
            }
            return text;
        }
    }
}
