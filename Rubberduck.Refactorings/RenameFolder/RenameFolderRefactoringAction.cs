using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.ChangeFolder;

namespace Rubberduck.Refactorings.RenameFolder
{
    public class RenameFolderRefactoringAction : CodeOnlyRefactoringActionBase<RenameFolderModel>
    {
        private readonly ICodeOnlyRefactoringAction<ChangeFolderModel> _changeFolder;

        public RenameFolderRefactoringAction(
            IRewritingManager rewritingManager,
            ChangeFolderRefactoringAction changeFolder)
            : base(rewritingManager)
        {
            _changeFolder = changeFolder;
        }

        public override void Refactor(RenameFolderModel model, IRewriteSession rewriteSession)
        {
            var sourceFolderParent = model.OriginalFolder.ParentFolder();
            var targetFolder = string.IsNullOrEmpty(sourceFolderParent)
                ? model.NewSubFolder
                : $"{sourceFolderParent}{FolderExtensions.FolderDelimiter}{model.NewSubFolder}";

            var changeFolderModel = new ChangeFolderModel(model.OriginalFolder, model.ModulesToMove, targetFolder);
            _changeFolder.Refactor(changeFolderModel, rewriteSession);
        }
    }
}