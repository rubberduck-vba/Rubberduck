using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting.Commands
{
    public class AddTestModuleWithStubsCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly AddTestModuleCommand _newUnitTestModuleCommand;

        public AddTestModuleWithStubsCommand(IVBE vbe, AddTestModuleCommand newUnitTestModuleCommand) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _newUnitTestModuleCommand = newUnitTestModuleCommand;
        }

        protected override bool EvaluateCanExecute(object parameter) => parameter is CodeExplorerComponentViewModel;

        protected override void OnExecute(object parameter)
        {
            if (parameter is CodeExplorerItemViewModel node)
            {
                _newUnitTestModuleCommand.Execute(node.Declaration);
            }
            else
            {
                _newUnitTestModuleCommand.Execute(_vbe.ActiveVBProject);
            }
        }
    }
}
