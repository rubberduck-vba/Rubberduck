using System;
using System.Collections.Generic;
namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        IDictionary<TestMethod, TestResult> AllTests { get; set; }
        IEnumerable<TestMethod> FailedTests();
        IEnumerable<TestMethod> LastRunTests(TestOutcome? outcome = null);
        IEnumerable<TestMethod> NotRunTests();
        IEnumerable<TestMethod> PassedTests();
        void ReRun();
        void Run();
        void Run(System.Collections.Generic.IEnumerable<TestMethod> tests);
        void RunFailedTests();
        void RunNotRunTests();
        void RunPassedTests();

        event EventHandler<TestCompleteEventArg> TestComplete;
        event EventHandler<EventArgs> AllTestsComplete;
    }

    public class TestCompleteEventArg : EventArgs
    {
        public TestResult Result { get; private set; }
        public TestMethod Test { get; private set; }

        public TestCompleteEventArg(TestMethod test, TestResult result)
        {
            this.Test = test;
            this.Result = result;
        }
    }
}
