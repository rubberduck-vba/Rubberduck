using Rubberduck.Parsing.Grammar;
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
            }
        }


        public string DeclarationStatement { get; set; } = string.Empty;

        public MoveCloserToUsageModel(Declaration target)
        {
            Target = target;
        }

    }
}