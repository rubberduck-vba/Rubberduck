using Rubberduck.UnitTesting;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that runs all Rubberduck unit tests in the VBE.
    /// </summary>
    public class RunAllTestsCommand : CommandBase
    {
        private readonly ITestEngine _engine;

        public RunAllTestsCommand(ITestEngine engine)
        {
            _engine = engine;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _engine.CanRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (_engine.CanRun)
            {
                _engine.Run(_engine.Tests);
            }
        }
    }
}
