using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageModel : IRefactoringModel
    {
        public Declaration Target { get; }

        public MoveCloserToUsageModel(Declaration target)
        {
            Target = target;
        }
    }
}