using NLog;
using Rubberduck.UI.Command;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI.UnitTesting.Commands
{
    class RunNotExecutedTestsCommand : CommandBase
    {
        private readonly ITestEngine testEngine;

        public RunNotExecutedTestsCommand(ITestEngine testEngine) : base(LogManager.GetCurrentClassLogger())
        {
            this.testEngine = testEngine;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return testEngine.CanRun;
        }

        protected override void OnExecute(object parameter)
        {
            if (!EvaluateCanExecute(parameter))
            {
                return;
            }
            testEngine.RunByOutcome(TestOutcome.Unknown);
        }
    }
}