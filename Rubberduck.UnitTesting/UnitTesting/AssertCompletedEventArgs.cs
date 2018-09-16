using System;

namespace Rubberduck.UnitTesting
{
    public class AssertCompletedEventArgs : EventArgs
    {
        public AssertCompletedEventArgs(TestOutcome outcome, string message = "")
        {
            Outcome = outcome;
            Message = message;
        }
        public TestOutcome Outcome { get; }
        public string Message { get; }
    }
}
