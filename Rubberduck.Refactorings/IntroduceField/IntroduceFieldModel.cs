using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.IntroduceField
{
    public class IntroduceFieldModel : IRefactoringModel
    {
        public Declaration Target { get; }

        public IntroduceFieldModel(Declaration target)
        {
            Target = target;
        }
    }
}