using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    [CodeExplorerCommand]
    public class AddStdModuleCommand : CommandBase
    {
        private readonly AddComponentCommand _addComponentCommand;

        public AddStdModuleCommand(AddComponentCommand addComponentCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _addComponentCommand = addComponentCommand;
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            return _addComponentCommand.CanAddComponent(parameter as CodeExplorerItemViewModel);
        }

        protected override void ExecuteImpl(object parameter)
        {
            _addComponentCommand.AddComponent(parameter as CodeExplorerItemViewModel, ComponentType.StandardModule);
        }
    }
}
