using System.Linq;
using System.Windows.Input;
using System.Windows.Threading;
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
        private readonly TestExplorerModelBase _model;

        public RunAllTestsCommand(ITestEngine engine, TestExplorerModelBase model)
        {
            _engine = engine;
            _model = model;
        }

        public override void Execute(object parameter)
        {
            _model.Refresh();
            _engine.Run(_model.Tests);
        }
    }
}