using System;

namespace Rubberduck.UnitTesting
{
    public class AssertCompletedEventArgs : EventArgs
    {
        public AssertCompletedEventArgs(TestOutcome outcome, string message = "")
        {
            _outcome = outcome;
            _message = message;
        }

        private readonly TestOutcome _outcome;
        public TestOutcome Outcome { get { return _outcome; } }

        private readonly string _message;
        public string Message { get { return _message; } }
    }
}
