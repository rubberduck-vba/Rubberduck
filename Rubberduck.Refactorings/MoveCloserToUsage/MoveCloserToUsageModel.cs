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
                DeclarationStatement = GetDefaultDeclarationStatement();
            }
        }


        public string DeclarationStatement { get; set; } = string.Empty;

        public MoveCloserToUsageModel(Declaration target)
        {
            Target = target;
        }

        private string GetDefaultDeclarationStatement()
        {
            /*
             * ToDo:  Use a less dirty Method to determine the original Variable Declaration ("Static" or "Dim" )
            */
            var completeDeclaration = Target.Context.Parent.Parent.GetText();

            if (Target.ParentDeclaration is ModuleDeclaration)
            {
                return Tokens.Static;                
            }
            else if (completeDeclaration.StartsWith(Tokens.Dim) )
            {
                return Tokens.Dim;
            }
            else if (completeDeclaration.StartsWith(Tokens.Static))
            {
                return Tokens.Static;
            }
            else
            {
                return Tokens.Dim;
            }
                

        }
    }
}