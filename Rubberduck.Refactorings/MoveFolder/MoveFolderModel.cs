using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveFolderModel : IRefactoringModel
    {
        public string SourceFolder { get; }
        public ICollection<ModuleDeclaration> ContainedModules { get; }
        public string TargetFolder { get; set; }

        public MoveFolderModel(string sourceFolder, ICollection<ModuleDeclaration> containedModules, string targetFolder)
        {
            SourceFolder = sourceFolder;
            ContainedModules = containedModules;
            TargetFolder = targetFolder;
        }
    }
}