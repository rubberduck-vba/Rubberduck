using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageModel : IRefactoringModel
    {
        private Declaration _target;
        public Declaration Target
        {
            get => _target;
            set
            {
                _target = value;
                NewDeclarationStatement = Parsing.Grammar.Tokens.Static;    // Static as Default to not affect semantics
            }
        }


        public string NewDeclarationStatement { get; set; } 

        public MoveCloserToUsageModel(Declaration target)
        {
            Target = target;
        }
    }
}