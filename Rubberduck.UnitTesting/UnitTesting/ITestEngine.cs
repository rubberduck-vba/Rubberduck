using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        IEnumerable<TestMethod> Tests { get; }
        void RunAll();
        void Run(IEnumerable<TestMethod> tests);
        event EventHandler<TestCompletedEventArgs> TestCompleted;
        event EventHandler TestsRefreshed;
        TestOutcome RunAggregateOutcome { get; }
        bool CanRun();
    }

    public class TestModuleEventArgs : EventArgs
    {
        public TestModuleEventArgs(QualifiedModuleName qualifiedModuleName)
        {
            QualifiedModuleName = qualifiedModuleName;
        }
        public QualifiedModuleName QualifiedModuleName { get; }
    }

    public class TestCompletedEventArgs : EventArgs
    {
        public TestMethod Test { get; private set; }
        public TestResult Result { get; private set; }

        public TestCompletedEventArgs(TestMethod test, TestResult result)
        {
            Test = test;
            Result = result;
        }
    }
}
