using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting.Commands
{
    internal class RunSucceededTestsCommand : CommandBase
    {
        private readonly ITestEngine _testEngine;

        public RunSucceededTestsCommand(ITestEngine testEngine)
        {
            _testEngine = testEngine;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _testEngine.CanRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }
            _testEngine.RunByOutcome(TestOutcome.Succeeded);
        }
    }
}
