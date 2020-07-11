using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ChangeFolder
{
    public class ChangeFolderModel : IRefactoringModel
    {
        public string OriginalFolder { get; }
        public ICollection<ModuleDeclaration> ModulesToMove { get; }
        public string NewFolder { get; set; }

        public ChangeFolderModel(string originalFolder, ICollection<ModuleDeclaration> modulesToMove, string newFolder)
        {
            OriginalFolder = originalFolder;
            ModulesToMove = modulesToMove;
            NewFolder = newFolder;
        }
    }
}