using System.Runtime.InteropServices;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodCommand : CommandBase
    {
        private readonly TestExplorerModel _model;
        private readonly NewTestMethodCommand _command;

        public AddTestMethodCommand(TestExplorerModel model, NewTestMethodCommand command)
        {
            _model = model;
            _command = command;
        }

        public override void Execute(object parameter)
        {
            _command.NewTestMethod();
        }
    }
}