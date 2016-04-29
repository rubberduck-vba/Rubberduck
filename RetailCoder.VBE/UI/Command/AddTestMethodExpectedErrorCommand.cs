using System.Runtime.InteropServices;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodExpectedErrorCommand : CommandBase
    {
        private readonly TestExplorerModelBase _model;
        private readonly NewTestMethodCommand _command;

        public AddTestMethodExpectedErrorCommand(TestExplorerModelBase model, NewTestMethodCommand command)
        {
            _model = model;
            _command = command;
        }

        public override void Execute(object parameter)
        {
            // legacy static class...
            var test = _command.NewExpectedErrorTestMethod();
            if (test != null)
            {
                _model.Tests.Add(test);
            }
        }
    }
}