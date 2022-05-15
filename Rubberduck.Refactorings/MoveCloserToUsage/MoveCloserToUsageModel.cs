using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageModel : IRefactoringModel
    {
        private VariableDeclaration _target;
        public VariableDeclaration Target
        {
            get => _target;
            set
            {
                _target = value;
            }
        }

        public string DeclarationStatement { get; set; } = string.Empty;

        public MoveCloserToUsageModel(VariableDeclaration target)
        {
            _target = target;
        }

        public MoveCloserToUsageModel(VariableDeclaration target, string declarationStatement )
        {
            _target = target;
            DeclarationStatement = declarationStatement;
        }

    }
}