namespace Rubberduck.UnitTesting
{
    public class TestResult
    {
        public void SetValues(TestOutcome outcome, string output = "", long duration = 0)
        {
            _outcome = outcome;
            _output = output;
            _duration = duration;
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
        }

        private long _duration;
        public long Duration { get { return _duration; } }

        private TestOutcome _outcome;
        public TestOutcome Outcome { get { return _outcome; } }

        private string _output;
        public string Output { get { return _output; } }
    }
}
