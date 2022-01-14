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
                DeclarationStatement = DefaultDeclarationStatement();
            }
        }


        public string DeclarationStatement { get; set; } = string.Empty;

        public MoveCloserToUsageModel(Declaration target)
        {
            Target = target;
        }

        private string DefaultDeclarationStatement()
        {            
            if (Target.ParentDeclaration is ModuleDeclaration)
            {
                return Tokens.Static;
            }
            else
            {
                //return Target.DeclarationType.ToString();
                // Get already used declaration Token (Static/Dim)
                return Tokens.Dim;
            }
                

        }
    }
}