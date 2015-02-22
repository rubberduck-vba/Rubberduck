using System;
using System.Collections.Generic;
namespace Rubberduck.UnitTesting
{
    interface ITestEngine
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

        event EventHandler<EventArgs> TestComplete;
        event EventHandler<EventArgs> AllTestsComplete;
    }
}
