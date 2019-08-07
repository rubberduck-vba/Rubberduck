using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting.Commands
{
    class RunFailedTestsCommand : CommandBase
    {
        private readonly ITestEngine testEngine;

        public RunFailedTestsCommand(ITestEngine testEngine)
        {
            this.testEngine = testEngine;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return testEngine.CanRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }
            testEngine.RunByOutcome(TestOutcome.Failed);
        }
    }
}
