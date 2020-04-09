using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.MoveMember;

namespace Rubberduck.Refactorings
{
    public interface IMoveMemberRefactoringPreviewer
    {
        string PreviewMove(MoveMemberModel model);
    }

    public interface IMoveMemberRefactoringPreviewerFactory
    {
        IMoveMemberRefactoringPreviewer Create(IMoveMemberEndpoint module);
    }

    public class MoveMemberRefactoringPreviewerFactory : IMoveMemberRefactoringPreviewerFactory
    {
        private readonly MoveMemberExistingModulesRefactoringAction _refactoringAction;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMovedContentProviderFactory _movedContentProviderFactory;

        public MoveMemberRefactoringPreviewerFactory(MoveMemberExistingModulesRefactoringAction refactoringAction, 
                                                IRewritingManager rewritingManager,
                                                IMovedContentProviderFactory movedContentProviderFactory)
        {
            _refactoringAction = refactoringAction;
            _rewritingManager = rewritingManager;
            _movedContentProviderFactory = movedContentProviderFactory;
        }

        public IMoveMemberRefactoringPreviewer Create(IMoveMemberEndpoint module)
        {
            return module is IMoveDestinationEndpoint
                ? new MoveMemberRefactoringDestinationPreviewer(_rewritingManager, _movedContentProviderFactory) as IMoveMemberRefactoringPreviewer
                : new MoveMemberRefactoringSourcePreviewer(_rewritingManager, _movedContentProviderFactory) as IMoveMemberRefactoringPreviewer;
        }
    }
}
