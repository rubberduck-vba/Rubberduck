using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.ChangeFolder;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveFolderRefactoringAction : CodeOnlyRefactoringActionBase<MoveFolderModel>
    {
        private readonly ICodeOnlyRefactoringAction<ChangeFolderModel> _changeFolder;

        public MoveFolderRefactoringAction(
            IRewritingManager rewritingManager,
            ChangeFolderRefactoringAction changeFolder)
            : base(rewritingManager)
        {
            _changeFolder = changeFolder;
        }

        public override void Refactor(MoveFolderModel model, IRewriteSession rewriteSession)
        {
            var sourceFolderParent = model.FolderToMove.ParentFolder();
            var changeFolderModel = new ChangeFolderModel(sourceFolderParent, model.ModulesToMove, model.TargetFolder);
            _changeFolder.Refactor(changeFolderModel, rewriteSession);
        }
    }
}