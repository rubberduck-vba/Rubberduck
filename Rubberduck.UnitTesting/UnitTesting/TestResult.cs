using System;

namespace Rubberduck.UnitTesting
{
    public class TestResult
    {
        public TestResult(TestOutcome outcome, string output = "", long duration = 0)
        {
            Outcome = outcome;
            Output = output;
            Duration = duration;
        }

        public long Duration { get; private set; }
        public TestOutcome Outcome { get; private set; }
        public string Output { get; private set; }
    }
}
