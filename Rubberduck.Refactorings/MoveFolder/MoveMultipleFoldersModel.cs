using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveFolder
{
    public class MoveMultipleFoldersModel : IRefactoringModel
    {
        public IDictionary<string, ICollection<ModuleDeclaration>> ModulesBySourceFolder { get; }
        public string TargetFolder { get; set; }

        public MoveMultipleFoldersModel(IDictionary<string, ICollection<ModuleDeclaration>> modulesBySourceFolder, string targetFolder)
        {
            ModulesBySourceFolder = modulesBySourceFolder;
            TargetFolder = targetFolder;
        }
    }
}