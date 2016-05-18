using System.Runtime.InteropServices;
using Rubberduck.Parsing.VBA;
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
        private readonly RubberduckParserState _state;

        public AddTestMethodExpectedErrorCommand(RubberduckParserState state, NewTestMethodCommand command)
        {
            _command = command;
            _state = state;
        }

        public override bool CanExecute(object parameter)
        {
            return _state.Status == ParserState.Ready;
        }

        public override void Execute(object parameter)
        {
            _command.NewExpectedErrorTestMethod();
        }
    }
}
