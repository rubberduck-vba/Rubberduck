using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class ImplementInterfaceModel : IRefactoringModel
    {
        public ClassModuleDeclaration TargetInterface { get; }
        public ClassModuleDeclaration TargetClass { get; }

        public ImplementInterfaceModel(ClassModuleDeclaration targetInterface, ClassModuleDeclaration targetClass)
        {
            TargetInterface = targetInterface;
            TargetClass = targetClass;
        }
    }
}