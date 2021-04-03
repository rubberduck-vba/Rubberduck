using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ImplicitTypeToExplicit
{
    public class ImplicitTypeToExplicitModel : IRefactoringModel
    {
        public ImplicitTypeToExplicitModel(Declaration target)
        {
            Target = target;
        }

        public Declaration Target { get; }

        public bool ForceVariantAsType { set; get; }
    }
}
