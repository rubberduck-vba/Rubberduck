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

        public string DeclarationStatement { get; set; }

        public MoveCloserToUsageModel(VariableDeclaration target, string declarationStatement = null )
        {
            _target = target;
            DeclarationStatement = declarationStatement ?? string.Empty;
        }

    }
}