using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveMultipleFoldersRefactoringAction : CodeOnlyRefactoringActionBase<MoveMultipleFoldersModel>
    {
        private readonly ICodeOnlyRefactoringAction<MoveFolderModel> _moveFolder;

        public MoveMultipleFoldersRefactoringAction(
            IRewritingManager rewritingManager,
            MoveFolderRefactoringAction moveFolder)
            : base(rewritingManager)
        {
            _moveFolder = moveFolder;
        }

        public override void Refactor(MoveMultipleFoldersModel model, IRewriteSession rewriteSession)
        {
            foreach (var sourceFolder in model.ModulesBySourceFolder.Keys)
            {
                var targetModel = new MoveFolderModel(sourceFolder, model.ModulesBySourceFolder[sourceFolder], model.TargetFolder);
                _moveFolder.Refactor(targetModel, rewriteSession);
            }
        }
    }
}