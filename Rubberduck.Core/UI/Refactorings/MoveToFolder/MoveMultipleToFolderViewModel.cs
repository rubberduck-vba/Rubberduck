using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Refactorings;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources;
using Rubberduck.Refactorings.MoveToFolder;
using Rubberduck.CodeAnalysis;

namespace Rubberduck.UI.Refactorings.MoveToFolder
{
    public class MoveMultipleToFolderViewModel : RefactoringViewModelBase<MoveMultipleToFolderModel>
    {
        public MoveMultipleToFolderViewModel(MoveMultipleToFolderModel model) 
            : base(model)
        {}

        private ICollection<ModuleDeclaration> Targets => Model.Targets;

        public string Instructions
        {
            get
            {
                if (Targets == null || !Targets.Any())
                {
                    return string.Empty;
                }

                if (Targets.Count == 1)
                {
                    var target = Targets.First();
                    var moduleName = target.IdentifierName;
                    var declarationType = CodeAnalysisUI.ResourceManager.GetString("DeclarationType_" + target.DeclarationType, CultureInfo.CurrentUICulture);
                    var currentFolder = target.CustomFolder;
                    return string.Format(RefactoringsUI.MoveToFolderDialog_InstructionsLabelText, declarationType, moduleName, currentFolder);
                }

                return string.Format(RefactoringsUI.MoveMultipleToFolderDialog_InstructionsLabelText);
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
                errors.Add(RefactoringsUI.MoveFolders_EmptyFolderName);
            }
            else
            {
                if (NewFolder.Any(char.IsControl))
                {
                    errors.Add(RefactoringsUI.MoveFolders_ControlCharacter);
                }

                if (NewFolder.Split(FolderExtensions.FolderDelimiter).Any(string.IsNullOrEmpty))
                {
                    errors.Add(RefactoringsUI.MoveFolders_EmptySubfolderName);
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

        public bool IsValidFolder => Targets != null 
                                     && Targets.Any()
                                     && !HasErrors;

        protected override void DialogOk()
        {
            if (Targets == null || !Targets.Any())
            {
                base.DialogCancel();
            }
            else
            {
                base.DialogOk();
            }
        }
    }
}
