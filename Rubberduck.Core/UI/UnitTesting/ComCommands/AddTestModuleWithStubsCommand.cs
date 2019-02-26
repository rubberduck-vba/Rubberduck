using NLog;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Command.ComCommands;
using Rubberduck.UI.UnitTesting.Commands;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.UnitTesting.ComCommands
{
    public class AddTestModuleWithStubsCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly AddTestModuleCommand _newUnitTestModuleCommand;

        public AddTestModuleWithStubsCommand(
            IVBE vbe, 
            AddTestModuleCommand newUnitTestModuleCommand,
            IVBEEvents vbeEvents) : base(LogManager.GetCurrentClassLogger(), vbeEvents)
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
