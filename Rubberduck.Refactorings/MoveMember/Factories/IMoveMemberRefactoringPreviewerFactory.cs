using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.MoveMember;

namespace Rubberduck.Refactorings
{
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
            if (module is IMoveDestinationEndpoint destination)
            {
                if (destination.IsExistingModule(out var destinationModule))
                {
                    return new MoveMemberRefactoringPreviewerDestination(_refactoringAction, _rewritingManager, _movedContentProviderFactory);
                }
                return new MoveMemberNullDestinationPreviewer(_refactoringAction, _rewritingManager, _movedContentProviderFactory);
            }
            return new MoveMemberRefactoringPreviewerSource(_refactoringAction, _rewritingManager);
        }
    }
}
