using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveToFolder
{
    public class MoveMultipleToFolderModel : IRefactoringModel
    {
        public ICollection<ModuleDeclaration> Targets { get; }
        public string TargetFolder { get; set; }

        public MoveMultipleToFolderModel(ICollection<ModuleDeclaration> targets, string targetFolder)
        {
            Targets = targets;
            TargetFolder = targetFolder;
        }
    }
}