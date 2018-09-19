using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddUserFormCommand : CommandBase
    {
        private readonly AddComponentCommand _addComponentCommand;

        public AddUserFormCommand(AddComponentCommand addComponentCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _addComponentCommand = addComponentCommand;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _addComponentCommand.CanAddComponent(parameter as CodeExplorerItemViewModel, ProjectTypes.VBA);
        }

        protected override void OnExecute(object parameter)
        {
            _addComponentCommand.AddComponent(parameter as CodeExplorerItemViewModel, ComponentType.UserForm);
        }
    }
}
