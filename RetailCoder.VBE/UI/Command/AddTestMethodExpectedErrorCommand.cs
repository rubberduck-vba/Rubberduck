using System.Runtime.InteropServices;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodExpectedErrorCommand : CommandBase
    {
        private readonly NewTestMethodCommand _command;

        public AddTestMethodExpectedErrorCommand(NewTestMethodCommand command)
        {
            _command = command;
        }

        public override void Execute(object parameter)
        {
            _command.NewExpectedErrorTestMethod();
        }
    }
}
