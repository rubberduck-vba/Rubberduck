using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all Rubberduck unit tests in the VBE.
    /// </summary>
    public class RunAllTestsCommand : CommandBase
    {
        private readonly ITestEngine _engine;
        private readonly TestExplorerModel _model;
        private readonly RubberduckParserState _state;

        public RunAllTestsCommand(RubberduckParserState state, ITestEngine engine, TestExplorerModel model)
        {
            _engine = engine;
            _model = model;
            _state = state;
        }

        public override void Execute(object parameter)
        {
            _state.StateChanged += StateChanged;

            _model.Refresh();
        }

        private void StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready) { return; }

            _model.ClearLastRun();
            _model.IsBusy = true;
            _engine.Run(_model.Tests);
            _model.IsBusy = false;
            _state.StateChanged -= StateChanged;
        }
    }
}
