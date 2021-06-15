using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.PromoteToParameter
{
    public class PromoteToParameterModel : IRefactoringModel
    {
        public Declaration Target { get; }
        public ModuleBodyElementDeclaration EnclosingMember { get; }

        public PromoteToParameterModel(Declaration target, ModuleBodyElementDeclaration enclosingMember)
        {
            Target = target;
            EnclosingMember = enclosingMember;
        }
    }
}