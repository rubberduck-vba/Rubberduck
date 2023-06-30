using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor.Events;

namespace Rubberduck.UI.Command.ComCommands
{
    public class PeekDefinitionFindAllReferencesCommand : ComCommandBase
    {
        private readonly FindAllReferencesAction _action;

        public PeekDefinitionFindAllReferencesCommand(
            FindAllReferencesAction action,
            IVbeEvents vbeEvents) : base(vbeEvents)
        {
            AddToCanExecuteEvaluation(EvaluateCanExecute, true);
            _action = action;
        }

        private bool EvaluateCanExecute(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration.IsUserDefined && !(declaration is ProjectDeclaration);
            }

            return false;
        }

        protected override void OnExecute(object parameter) => _action?.FindAllReferences((Declaration)parameter);
    }

    public class PeekDefinitionNavigateCommand : ComCommandBase
    {
        private readonly INavigateCommand _action;

        public PeekDefinitionNavigateCommand(INavigateCommand action, IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            AddToCanExecuteEvaluation(EvaluateCanExecute, true);
            _action = action;
        }

        private bool EvaluateCanExecute(object parameter)
        {
            if (parameter is Declaration declaration)
            {
                return declaration.IsUserDefined
                       && !(declaration is ProjectDeclaration);
            }

            return false;
        }

        protected override void OnExecute(object parameter) =>
            _action?.Execute(((Declaration) parameter).QualifiedSelection.GetNavitationArgs());
    }
}