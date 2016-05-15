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

        public RunAllTestsCommand(ITestEngine engine, TestExplorerModel model)
        {
            _engine = engine;
            _model = model;
        }

        public override void Execute(object parameter)
        {
            _model.Refresh();
            _model.ClearLastRun();
            _model.IsBusy = true;
            _engine.Run(_model.Tests);
            _model.IsBusy = false;
        }
    }
}