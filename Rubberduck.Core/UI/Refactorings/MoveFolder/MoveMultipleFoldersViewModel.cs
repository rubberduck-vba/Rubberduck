using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveFolder;
using Rubberduck.Resources;
using Rubberduck.JunkDrawer.Extensions;

namespace Rubberduck.UI.Refactorings.MoveFolder
{
    public class MoveMultipleFoldersViewModel : RefactoringViewModelBase<MoveMultipleFoldersModel>
    {
        public MoveMultipleFoldersViewModel(MoveMultipleFoldersModel model) 
            : base(model)
        {}

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
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsValidFolder));
            }
        }
        
        public bool IsValidFolder => ModulesBySourceFolder != null 
                                     && ModulesBySourceFolder.Any()
                                     && NewFolder != null 
                                     && !NewFolder.Any(char.IsControl);
        
        protected override void DialogOk()
        {
            if (ModulesBySourceFolder == null || !ModulesBySourceFolder.Any())
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
