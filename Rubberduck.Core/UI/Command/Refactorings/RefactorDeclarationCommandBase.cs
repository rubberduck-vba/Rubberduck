using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;

namespace Rubberduck.UI.Command.Refactorings
{
    public abstract class RefactorDeclarationCommandBase : RefactorCommandBase
    {
        protected RefactorDeclarationCommandBase(IRefactoring refactoring, IParserStatusProvider parserStatusProvider) 
            : base(refactoring, parserStatusProvider)
        {}

        protected override void OnExecute(object parameter)
        {
            var target = GetTarget();
            if (target != null)
            {
                Refactoring.Refactor(target);
            }
        }

        protected abstract Declaration GetTarget();
    }
}