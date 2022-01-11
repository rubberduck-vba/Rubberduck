using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageModel : IRefactoringModel
    {
        public Declaration Target { get; }

        public string NewDeclarationStatement { get; set; } = Parsing.Grammar.Tokens.Static;

        public MoveCloserToUsageModel(Declaration target)
        {
            Target = target;
        }
    }
}