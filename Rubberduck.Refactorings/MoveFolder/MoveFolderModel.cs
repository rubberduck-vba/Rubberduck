using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveFolderModel : IRefactoringModel
    {
        public string FolderToMove { get; }
        public ICollection<ModuleDeclaration> ModulesToMove { get; }
        public string TargetFolder { get; set; }

        public MoveFolderModel(string folderToMove, ICollection<ModuleDeclaration> modulesToMove, string targetFolder)
        {
            FolderToMove = folderToMove;
            ModulesToMove = modulesToMove;
            TargetFolder = targetFolder;
        }
    }
}