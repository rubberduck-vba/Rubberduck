using System.Linq;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.MoveToFolder;

namespace Rubberduck.Refactorings.ChangeFolder
{
    public class ChangeFolderRefactoringAction : CodeOnlyRefactoringActionBase<ChangeFolderModel>
    {
        private readonly ICodeOnlyRefactoringAction<MoveToFolderModel> _moveToFolder;

        public ChangeFolderRefactoringAction(
            IRewritingManager rewritingManager,
            MoveToFolderRefactoringAction moveToFolder)
            : base(rewritingManager)
        {
            _moveToFolder = moveToFolder;
        }

        public override void Refactor(ChangeFolderModel model, IRewriteSession rewriteSession)
        {
            var originalFolder = model.OriginalFolder;

            foreach (var module in model.ModulesToMove.Distinct())
            {
                var currentFolder = module.CustomFolder;

                if (!currentFolder.StartsWith(originalFolder))
                {
                    continue;
                }
                
                var newFolder = currentFolder.Equals(originalFolder)
                    ? model.NewFolder
                    : $"{model.NewFolder}{FolderExtensions.FolderDelimiter}{currentFolder.SubFolderPathRelativeTo(model.OriginalFolder)}";

                var moduleModel = new MoveToFolderModel(module, newFolder);
                _moveToFolder.Refactor(moduleModel, rewriteSession);
            }
        }
    }
}