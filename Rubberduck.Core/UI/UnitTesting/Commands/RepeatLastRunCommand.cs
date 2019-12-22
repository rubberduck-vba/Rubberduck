using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting.Commands
{
    internal class RepeatLastRunCommand : CommandBase
    {
        private ITestEngine testEngine;
        public RepeatLastRunCommand(ITestEngine testEngine)
        {
            this.testEngine = testEngine;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return testEngine.CanRun && testEngine.CanRepeatLastRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }
            testEngine.RepeatLastRun();
        }
    }
}
