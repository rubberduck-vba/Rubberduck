using Rubberduck.Parsing.Rewriter;
using Rubberduck.VBEditor;
using System;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringPreviewer
    {
        string PreviewMove(MoveMemberModel model);
    }

    public abstract class MoveMemberRefactoringPreviewProvider : RefactoringPreviewProviderWrapperBase<MoveMemberModel>
    {
        public MoveMemberRefactoringPreviewProvider(MoveMemberExistingModulesRefactoringAction refactoringAction, 
            IRewritingManager rewritingManager)
             : base(refactoringAction, rewritingManager)
        {}

        protected abstract override QualifiedModuleName ComponentToShow(MoveMemberModel model);

        public static string FormatPreview(string text, int maxConsecutiveNewLines = 2, bool addDemarcation = false)
        {
            var target = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines + 1).ToList());
            var replacement = string.Concat(Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            for (var counter = 1; counter < 10 && text.Contains(target); counter++)
            {
                text = text.Replace(target, replacement);
            }
            return addDemarcation
                    ? $"'*****  {Resources.RubberduckUI.MoveMember_MovedContentBelowThisLine}  *****{Environment.NewLine}{Environment.NewLine}{text}{Environment.NewLine}{Environment.NewLine}'****  {Resources.RubberduckUI.MoveMember_MovedContentAboveThisLine}  ****{Environment.NewLine}"
                    : text;
        }
    }

    public class MoveMemberRefactoringPreviewerSource : MoveMemberRefactoringPreviewProvider, IMoveMemberRefactoringPreviewer
    {
        public MoveMemberRefactoringPreviewerSource(MoveMemberExistingModulesRefactoringAction refactoringAction, 
            IRewritingManager rewritingManager)
            :base(refactoringAction, rewritingManager)
        {
        }

        protected override QualifiedModuleName ComponentToShow(MoveMemberModel model)
        {
            return model.Source.QualifiedModuleName;
        }

        public string PreviewMove(MoveMemberModel model)
        {
            return FormatPreview(base.Preview(model));
        }
    }

    public class MoveMemberRefactoringPreviewerDestination : MoveMemberRefactoringPreviewProvider, IMoveMemberRefactoringPreviewer
    {
        private QualifiedModuleName _qmn;
        private readonly IMovedContentProviderFactory _movedContentProviderFactory;

        public MoveMemberRefactoringPreviewerDestination(MoveMemberExistingModulesRefactoringAction refactoringAction,
            IRewritingManager rewritingManager,
            IMovedContentProviderFactory movedContentProviderFactory)
            : base(refactoringAction, rewritingManager)
        {
            _movedContentProviderFactory = movedContentProviderFactory;
        }

        protected override QualifiedModuleName ComponentToShow(MoveMemberModel model)
        {
            if (!model.Destination.IsExistingModule(out var module))
            {
                throw new ArgumentException();
            }
            _qmn = module.QualifiedModuleName;
            return _qmn;
        }

        public string PreviewMove(MoveMemberModel model)
        {
            if (!model.SelectedDeclarations.Any())
            {
                return FormatPreview(Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove, addDemarcation: true);
            }

            if (!model.TryFindApplicableStrategy(out var strategy))
            {
                return FormatPreview(Resources.RubberduckUI.MoveMember_ApplicableStrategyNotFound, addDemarcation: true);
            }

            ((MoveMemberExistingModulesRefactoringAction)_refactoringAction).ContentProvider = _movedContentProviderFactory.CreatePreviewProvider();

            var preview = FormatPreview(base.Preview(model));

            return preview;
        }
    }

    public class MoveMemberNullDestinationPreviewer : IMoveMemberRefactoringPreviewer
    {
        private readonly MoveMemberExistingModulesRefactoringAction _refactoring;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMovedContentProviderFactory _movedContentProviderFactory;

        public MoveMemberNullDestinationPreviewer(
                        MoveMemberExistingModulesRefactoringAction refactoring,
                        IRewritingManager rewritingManager,
                        IMovedContentProviderFactory movedContentProviderFactory)
        {
            _refactoring = refactoring;
            _rewritingManager = rewritingManager;
            _movedContentProviderFactory = movedContentProviderFactory;
        }

        public string PreviewMove(MoveMemberModel model)
        {
            if (!model.TryFindApplicableStrategy(out var strategy))
            {
                return MoveMemberRefactoringPreviewProvider.FormatPreview(Resources.RubberduckUI.MoveMember_ApplicableStrategyNotFound, addDemarcation: true);
            }

            if (!model.SelectedDeclarations.Any())
            {
                return MoveMemberRefactoringPreviewProvider.FormatPreview(Resources.RubberduckUI.MoveMember_NoDeclarationsSelectedToMove, addDemarcation: true);
            }

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();

            _refactoring.ContentProvider = _movedContentProviderFactory.CreatePreviewProvider();

            var preview = _refactoring.NewModuleContent(model, rewriteSession);

            return MoveMemberRefactoringPreviewProvider.FormatPreview(preview.Trim());
        }
    }
}
