using System;
using System.Collections.Generic;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        event EventHandler<TestCompletedEventArgs> TestCompleted;
        event EventHandler<TestRunCompletedEventArgs> TestRunCompleted;
        event EventHandler TestsRefreshed;

        IEnumerable<TestMethod> Tests { get; }
        TestOutcome CurrentAggregateOutcome { get; }
        bool CanRun { get; }
        bool CanRepeatLastRun { get; }

        void Run(IEnumerable<TestMethod> tests);
        void RunByOutcome(TestOutcome outcome);
        void RepeatLastRun();
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
