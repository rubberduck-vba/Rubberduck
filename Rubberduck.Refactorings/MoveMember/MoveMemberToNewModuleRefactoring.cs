using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberToNewModuleRefactoring : RefactoringActionWithSuspension<MoveMemberModel>
    {
        private readonly MoveMemberToExistingModuleRefactoring _refactoring;
        private readonly IRewritingManager _rewritingManager;
        private readonly IAddComponentService _addComponentService;

        public MoveMemberToNewModuleRefactoring(
                        MoveMemberToExistingModuleRefactoring refactoring,
                        IParseManager parseManager,
                        IRewritingManager rewritingManager,
                        IAddComponentService addComponentService)
                : base(parseManager, rewritingManager)
        {
            _refactoring = refactoring;
            _rewritingManager = rewritingManager;
            _addComponentService = addComponentService;
        }

        protected override void Refactor(MoveMemberModel model, IRewriteSession rewriteSession)
        {
            if (!MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy) 
                || !strategy.IsExecutableModel(model, out _))
            {
                return;
            }

            var newContent = strategy.NewDestinationModuleContent(model, _rewritingManager, new MovedContentProvider()).AsSingleBlock;

            _refactoring.Refactor(model, rewriteSession);

            _addComponentService.AddComponentWithAttributes(
                                        model.Source.Module.ProjectId,
                                        model.Destination.ComponentType,
                                        newContent,
                                        componentName: model.Destination.ModuleName);
        }

        protected override bool RequiresSuspension(MoveMemberModel model) => true;
    }
}
