using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.IntroduceParameter
{
    public class IntroduceParameterModel : IRefactoringModel
    {
        public Declaration Target { get; }
        public ModuleBodyElementDeclaration EnclosingMember { get; }

        public IntroduceParameterModel(Declaration target, ModuleBodyElementDeclaration enclosingMember)
        {
            Target = target;
            EnclosingMember = enclosingMember;
        }
    }
}