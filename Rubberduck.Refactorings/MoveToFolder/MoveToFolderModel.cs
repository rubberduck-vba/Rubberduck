using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveToFolder
{
    public class MoveToFolderModel : IRefactoringModel
    {
        public ModuleDeclaration Target { get; }
        public string TargetFolder { get; set; }

        public MoveToFolderModel(ModuleDeclaration target, string targetFolder)
        {
            Target = target;
            TargetFolder = targetFolder;
        }
    }
}