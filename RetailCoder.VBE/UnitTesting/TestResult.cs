using System;

namespace Rubberduck.UnitTesting
{
    public class TestResult
    {
        public void SetValues(TestOutcome outcome, string output = "", long duration = 0, DateTime? startTime = null, DateTime? endTime = null)
        {
            _outcome = outcome;
            _output = output;
            _duration = duration;
            _startTime = startTime ?? DateTime.Now;
            _endTime = endTime ?? DateTime.Now;
        }

        public void SetDuration(long duration)
        {
            _duration = duration;
        }

        public TestResult(TestOutcome outcome, string output = "", long duration = 0)
        {
            _outcome = outcome;
            _output = output;
            _duration = duration;
            _startTime = DateTime.Now;
            _endTime = DateTime.Now;
        }

        private long _duration;
        public long Duration { get { return _duration; } }

        private DateTime _startTime;
        public DateTime StartTime { get { return _startTime; } }

        private DateTime _endTime;
        public DateTime EndTime { get { return _endTime; } }

        private TestOutcome _outcome;
        public TestOutcome Outcome { get { return _outcome; } }

        private string _output;
        public string Output { get { return _output; } }
    }
}
