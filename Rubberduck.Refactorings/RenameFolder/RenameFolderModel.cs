using System.Collections.Generic;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.RenameFolder
{
    public class RenameFolderModel : IRefactoringModel
    {
        public string OriginalFolder { get; }
        public ICollection<ModuleDeclaration> ModulesToMove { get; }
        public string NewSubFolder { get; set; }

        public RenameFolderModel(string originalFolder, ICollection<ModuleDeclaration> modulesToMove, string newSubFolder)
        {
            OriginalFolder = originalFolder;
            ModulesToMove = modulesToMove;
            NewSubFolder = newSubFolder;
        }

        public string SubFolderToRename => OriginalFolder.SubFolderName();
    }
}