using NLog;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting.Commands
{
    internal class RepeatLastRunCommand : CommandBase
    {
        private ITestEngine testEngine;
        public RepeatLastRunCommand(ITestEngine testEngine) : base (LogManager.GetCurrentClassLogger())
        {
            this.testEngine = testEngine;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return testEngine.CanRun && testEngine.CanRepeatLastRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (!EvaluateCanExecute(parameter))
            {
                return;
            }
            testEngine.RepeatLastRun();
        }
    }
}
