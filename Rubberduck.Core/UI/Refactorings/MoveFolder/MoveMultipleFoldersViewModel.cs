using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Resources;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Refactorings.MoveFolder
{
    public class MoveMultipleFoldersViewModel : RefactoringViewModelBase<MoveMultipleFoldersModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;

        public MoveMultipleFoldersViewModel(MoveMultipleFoldersModel model, IMessageBox messageBox, IDeclarationFinderProvider declarationFinderProvider)
            : base(model)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _messageBox = messageBox;
        }

        private IDictionary<string, ICollection<ModuleDeclaration>> ModulesBySourceFolder => Model.ModulesBySourceFolder;

        public string Instructions
        {
            get
            {
                if (ModulesBySourceFolder == null || !ModulesBySourceFolder.Any())
                {
                    return string.Empty;
                }

                var sourceFolders = ModulesBySourceFolder.Keys;
                if (sourceFolders.Count == 1)
                {
                    var sourceFolder = sourceFolders.First();
                    var sourceParent = sourceFolder.ParentFolder();

                    if (sourceParent.Length == 0)
                    {
                        return string.Format(RubberduckUI.MoveRootFolderDialog_InstructionsLabelText, sourceFolder);
                    }

                    return string.Format(RubberduckUI.MoveFolderDialog_InstructionsLabelText, sourceFolder, sourceParent);
                }

                return string.Format(RubberduckUI.MoveFoldersDialog_InstructionsLabelText);
            }
        }

        public string NewFolder
        {
            get => Model.TargetFolder;
            set
            {
                Model.TargetFolder = value;
                ValidateFolder();
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidFolder));
            }
        }

        private void ValidateFolder()
        {
            var errors = new List<string>();

            if (string.IsNullOrEmpty(NewFolder))
            {
                errors.Add(RubberduckUI.MoveFolders_EmptyFolderName);
            }
            else
            {
                if (NewFolder.Any(char.IsControl))
                {
                    errors.Add(RubberduckUI.MoveFolders_ControlCharacter);
                }

                if (NewFolder.Split(FolderExtensions.FolderDelimiter).Any(string.IsNullOrEmpty))
                {
                    errors.Add(RubberduckUI.MoveFolders_EmptySubfolderName);
                }
            }

            if (errors.Any())
            {
                SetErrors(nameof(NewFolder), errors);
            }
            else
            {
                ClearErrors();
            }
        }
        
        public bool IsValidFolder => ModulesBySourceFolder != null 
                                     && ModulesBySourceFolder.Any()
                                     && !HasErrors;
        
        protected override void DialogOk()
        {
            if (ModulesBySourceFolder == null 
                || !ModulesBySourceFolder.Any()
                || MergesSourceFolders() && !UserConfirmsToProceedWithSourceFolderMerge())
            {
                base.DialogCancel();
            }
            else
            {
                var foldersMergedWithTargetFolders = FoldersMergedWithTargetFolders();
                if (foldersMergedWithTargetFolders.Any()
                    && !UserConfirmsToProceedWithFolderMerge(NewFolder, foldersMergedWithTargetFolders))
                {
                    base.DialogCancel();
                }
                else
                {
                    base.DialogOk();
                }
            }
        }

        private bool MergesSourceFolders()
        {
            return ModulesBySourceFolder
                .Select(kvp => kvp.Key.SubFolderName())
                .GroupBy(item => item)
                .Any(group => group.Count() > 1);
        }

        private bool UserConfirmsToProceedWithSourceFolderMerge()
        {
            var message = RubberduckUI.MoveFolders_SameNameSourceFolders;
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.MoveFoldersDialog_Caption) ?? false;
        }

        private List<string> FoldersMergedWithTargetFolders()
        {
            var movingFolders = ModulesBySourceFolder
                .Select(kvp => kvp.Key)
                .Where(folder => !folder.ParentFolder().Equals(NewFolder))
                .Select(folder => folder.SubFolderName());

            var targetFolderSubfolders = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .Select(module => module.CustomFolder)
                .Where(folder => folder.IsSubFolderOf(NewFolder))
                .Select(folder => folder.SubFolderPathRelativeTo(NewFolder).RootFolder())
                .ToHashSet();

            return movingFolders
                .Where(folder => targetFolderSubfolders.Contains(folder))
                .ToList();
        }

        private bool UserConfirmsToProceedWithFolderMerge(string targetFolder, List<string> mergedTargetFolders)
        {
            var message = FolderMergeUserConfirmationMessage(targetFolder, mergedTargetFolders);
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.MoveFoldersDialog_Caption) ?? false;
        }

        private string FolderMergeUserConfirmationMessage(string targetFolder, List<string> mergedTargetFolders)
        {
            if (mergedTargetFolders.Count == 1)
            {
                return string.Format(
                    RubberduckUI.MoveFolders_SameNameSubfolder, 
                    targetFolder,
                    mergedTargetFolders.Single());
            }

            var subfolders = $"'{string.Join("', '", mergedTargetFolders)}'";
            return string.Format(
                RubberduckUI.MoveFolders_SameNameSubfolders,
                targetFolder,
                subfolders);
        }
    }
}
