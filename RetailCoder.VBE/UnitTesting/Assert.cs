using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    public enum TestOutcome
    {
        Unknown,
        Inconclusive,
        Failed,
        Succeeded
    }

    public class TestResult
    {
        public static TestResult Unknown()
        {
            return new TestResult(TestOutcome.Unknown);
        }

        public static TestResult Success()
        {
            return new TestResult(TestOutcome.Succeeded);
        }

        public static TestResult Inconclusive(string message = null)
        {
            return new TestResult(TestOutcome.Inconclusive, message);
        }

        public TestResult(TestOutcome outcome, string output = null)
        {
            _outcome = outcome;
            _output = output;
        }

        public TestResult(TestResult result, long elapsedMilliseconds)
            :this(result.Outcome, result.Output)
        {
            _elapsedMilliseconds = elapsedMilliseconds;
        }

        private readonly long _elapsedMilliseconds;
        public long Duration { get { return _elapsedMilliseconds; } }

        private readonly TestOutcome _outcome;
        public TestOutcome Outcome { get { return _outcome; } }

        private readonly string _output;
        public string Output { get { return _output; } }
    }

    [ComVisible(false)]
    internal class AssertCompletedEventArgs : EventArgs
    {
        public AssertCompletedEventArgs(TestResult result)
        {
            _result = result;
        }

        private readonly TestResult _result;
        public TestResult Result { get { return _result; } }
    }
}
