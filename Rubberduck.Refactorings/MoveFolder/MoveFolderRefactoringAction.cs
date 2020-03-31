using System.Linq;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings.MoveToFolder;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveFolderRefactoringAction : CodeOnlyRefactoringActionBase<MoveFolderModel>
    {
        private readonly ICodeOnlyRefactoringAction<MoveToFolderModel> _moveToFolder;

        public MoveFolderRefactoringAction(
            IRewritingManager rewritingManager,
            MoveToFolderRefactoringAction moveToFolder)
            : base(rewritingManager)
        {
            _moveToFolder = moveToFolder;
        }

        public override void Refactor(MoveFolderModel model, IRewriteSession rewriteSession)
        {
            var sourceFolderParent = model.SourceFolder.ParentFolder();

            foreach (var module in model.ContainedModules.Distinct())
            {
                var currentFolder = module.CustomFolder;
                var subPath = currentFolder.SubFolderPathRelativeTo(sourceFolderParent);
                var newFolder = $"{model.TargetFolder}{FolderExtensions.FolderDelimiter}{subPath}";
                var moduleModel = new MoveToFolderModel(module, newFolder);
                _moveToFolder.Refactor(moduleModel, rewriteSession);
            }
        }
    }
}