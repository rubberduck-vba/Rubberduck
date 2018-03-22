using System;

namespace Rubberduck.UnitTesting
{
    public class TestResult
    {
        public void SetValues(TestOutcome outcome, string output = "", long duration = 0, DateTime? startTime = null, DateTime? endTime = null)
        {
            Outcome = outcome;
            Output = output;
            Duration = duration;
            StartTime = startTime ?? DateTime.Now;
            EndTime = endTime ?? DateTime.Now;
        }

        public void SetDuration(long duration)
        {
            Duration = duration;
        }

        public TestResult(TestOutcome outcome, string output = "", long duration = 0)
        {
            Outcome = outcome;
            Output = output;
            Duration = duration;
            StartTime = DateTime.Now;
            EndTime = DateTime.Now;
        }

        public long Duration { get; private set; }

        public DateTime StartTime { get; private set; }

        public DateTime EndTime { get; private set; }

        public TestOutcome Outcome { get; private set; }

        public string Output { get; private set; }
    }
}
