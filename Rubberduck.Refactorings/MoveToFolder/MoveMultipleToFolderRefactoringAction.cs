using System.Linq;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Refactorings.MoveToFolder
{
    public class MoveMultipleToFolderRefactoringAction : CodeOnlyRefactoringActionBase<MoveMultipleToFolderModel>
    {
        private readonly ICodeOnlyRefactoringAction<MoveToFolderModel> _moveToFolder;

        public MoveMultipleToFolderRefactoringAction(
            IRewritingManager rewritingManager,
            MoveToFolderRefactoringAction moveToFolder)
            : base(rewritingManager)
        {
            _moveToFolder = moveToFolder;
        }

        public override void Refactor(MoveMultipleToFolderModel model, IRewriteSession rewriteSession)
        {
            foreach (var target in model.Targets.Distinct())
            {
                var targetModel = new MoveToFolderModel(target, model.TargetFolder);
                _moveToFolder.Refactor(targetModel, rewriteSession);
            }
        }
    }
}