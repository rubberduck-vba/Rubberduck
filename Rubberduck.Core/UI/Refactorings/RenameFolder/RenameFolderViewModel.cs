using System.Collections.Generic;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.RenameFolder;
using Rubberduck.Resources;

namespace Rubberduck.UI.Refactorings.RenameFolder
{
    public class RenameFolderViewModel : RefactoringViewModelBase<RenameFolderModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;

        public RenameFolderViewModel(RenameFolderModel model, IMessageBox messageBox, IDeclarationFinderProvider declarationFinderProvider)
            : base(model)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _messageBox = messageBox;
        }

        public string Instructions
        {
            get
            {
                if (string.IsNullOrEmpty(Model?.OriginalFolder))
                {
                    return string.Empty;
                }

                var folderToRename = Model.OriginalFolder;

                return string.Format(
                    RubberduckUI.RenameDialog_InstructionsLabelText,
                    RubberduckUI.RenameDialog_Folder,
                    folderToRename);
            }
        }

        public string NewFolderName
        {
            get => Model.NewSubFolder;
            set
            {
                Model.NewSubFolder = value;
                ValidateFolder();
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidFolder));
                OnPropertyChanged(nameof(FullNewFolderName));
            }
        }

        public string FullNewFolderName => Model.OriginalFolder.Contains(FolderExtensions.FolderDelimiter)
            ? $"{Model.OriginalFolder.ParentFolder()}{FolderExtensions.FolderDelimiter}{NewFolderName}"
            : NewFolderName;

        private void ValidateFolder()
        {
            var errors = new List<string>();

            if (string.IsNullOrEmpty(NewFolderName))
            {
                //We generally already rename a subfolder, here.
                errors.Add(RubberduckUI.MoveFolders_EmptySubfolderName);
            }
            else
            {
                if (NewFolderName.Any(char.IsControl))
                {
                    errors.Add(RubberduckUI.MoveFolders_ControlCharacter);
                }

                if (NewFolderName.Split(FolderExtensions.FolderDelimiter).Any(string.IsNullOrEmpty))
                {
                    errors.Add(RubberduckUI.MoveFolders_EmptySubfolderName);
                }
            }

            if (errors.Any())
            {
                SetErrors(nameof(NewFolderName), errors);
            }
            else
            {
                ClearErrors();
            }
        }

        public bool IsValidFolder => Model?.ModulesToMove != null
                                     && Model.ModulesToMove.Any()
                                     && !HasErrors;

        protected override void DialogOk()
        {
            if (Model?.ModulesToMove == null
                || !Model.ModulesToMove.Any()
                || !Model.SubFolderToRename.Equals(Model.NewSubFolder)
                    && FolderAlreadyExists(FullNewFolderName)
                    && !UserConfirmsToProceedWithFolderMerge(FullNewFolderName, Model.SubFolderToRename, NewFolderName))
            {
                base.DialogCancel();
            }
            else
            {
                base.DialogOk();
            }
        }

        private bool FolderAlreadyExists(string fullFolderName)
        {
            return _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Module)
                .OfType<ModuleDeclaration>()
                .Select(module => module.CustomFolder)
                .Any(folder => folder.Equals(fullFolderName) 
                               || folder.IsSubFolderOf(fullFolderName));
        }

        private bool UserConfirmsToProceedWithFolderMerge(string fullTargetFolder, string partToRename, string newFolderPart)
        {
            var message = string.Format(
                RubberduckUI.RenameDialog_FolderAlreadyExists,
                fullTargetFolder,
                partToRename,
                newFolderPart);
            return _messageBox?.ConfirmYesNo(message, RubberduckUI.RenameDialog_Caption) ?? false;
        }
    }
}