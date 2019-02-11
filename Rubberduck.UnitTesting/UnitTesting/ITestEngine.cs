using System;
using System.Collections.Generic;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        event EventHandler<TestRunStartedEventArgs> TestRunStarted;
        event EventHandler<TestStartedEventArgs> TestStarted;
        event EventHandler<TestCompletedEventArgs> TestCompleted;
        event EventHandler<TestRunCompletedEventArgs> TestRunCompleted;
        event EventHandler TestsRefreshStarted;
        event EventHandler TestsRefreshed;

        IEnumerable<TestMethod> Tests { get; }
        IReadOnlyList<TestMethod> LastRunTests { get; }
        bool CanRun { get; }
        bool CanRepeatLastRun { get; }
        void Run(IEnumerable<TestMethod> tests);
        void RunByOutcome(TestOutcome outcome);
        void RepeatLastRun();
        void RequestCancellation();
    }

    public class TestRunStartedEventArgs : EventArgs
    {
        public IReadOnlyList<TestMethod> Tests { get; }

        public TestRunStartedEventArgs(IReadOnlyList<TestMethod> tests)
        {
            Tests = tests;
        }
    }

    public class TestStartedEventArgs : EventArgs
    {
        public TestMethod Test { get; }

        public TestStartedEventArgs(TestMethod test)
        {
            Test = test;
        }
    }

    public class TestCompletedEventArgs : EventArgs
    {
        public TestMethod Test { get; }
        public TestResult Result { get; }

        public TestCompletedEventArgs(TestMethod test, TestResult result)
        {
            Test = test;
            Result = result;
        }
    }

    public class TestRunCompletedEventArgs : EventArgs
    {
        public long Duration { get; }

        public TestRunCompletedEventArgs(long duration)
        {
            Duration = duration;
        }
    }
}
