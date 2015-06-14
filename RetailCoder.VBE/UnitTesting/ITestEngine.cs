using System;
using System.Collections.Generic;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        IDictionary<TestMethod, TestResult> AllTests { get; set; }
        IEnumerable<TestMethod> FailedTests();
        IEnumerable<TestMethod> LastRunTests(TestOutcome? outcome = null);
        IEnumerable<TestMethod> NotRunTests();
        IEnumerable<TestMethod> PassedTests();

        event EventHandler<TestModuleEventArgs> ModuleInitialize;
        event EventHandler<TestModuleEventArgs> ModuleCleanup;
        event EventHandler<TestModuleEventArgs> MethodInitialize;
        event EventHandler<TestModuleEventArgs> MethodCleanup;
        void Run(IEnumerable<TestMethod> tests, VBProject vbe);

        event EventHandler<TestCompletedEventArgs> TestComplete;
    }

    public class TestModuleEventArgs : EventArgs
    {
        public TestModuleEventArgs(QualifiedModuleName qualifiedModuleName)
        {
            _qualifiedModuleName = qualifiedModuleName;
        }

        private readonly QualifiedModuleName _qualifiedModuleName;
        public QualifiedModuleName QualifiedModuleName { get { return _qualifiedModuleName; } }
    }

    public class TestMethodEventArgs : TestModuleEventArgs
    {
        public TestMethodEventArgs(QualifiedMemberName qualifiedMemberName)
            :base(qualifiedMemberName.QualifiedModuleName)
        {
            _qualifiedMemberName = qualifiedMemberName;
        }

        private readonly QualifiedMemberName _qualifiedMemberName;
        public QualifiedMemberName QualifiedMemberName { get { return _qualifiedMemberName; } }
    }

    public class TestCompletedEventArgs : EventArgs
    {
        public TestResult Result { get; private set; }
        public TestMethod Test { get; private set; }

        public TestCompletedEventArgs(TestMethod test, TestResult result)
        {
            Test = test;
            Result = result;
        }
    }
}
