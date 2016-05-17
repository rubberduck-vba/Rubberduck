namespace Rubberduck.UnitTesting
{
    public class TestResult
    {
        public void SetValues(TestOutcome outcome, string output = "", long elapsedMilliseconds = 0)
        {
            _outcome = outcome;
            _output = output;
            _elapsedMilliseconds = elapsedMilliseconds;
        }

        public TestResult(TestOutcome outcome, string output = "", long elapsedMilliseconds = 0)
        {
            _outcome = outcome;
            _output = output;
            _elapsedMilliseconds = elapsedMilliseconds;
        }

        private long _elapsedMilliseconds;
        public long Duration { get { return _elapsedMilliseconds; } }

        private TestOutcome _outcome;
        public TestOutcome Outcome { get { return _outcome; } }

        private string _output;
        public string Output { get { return _output; } }
    }
}
