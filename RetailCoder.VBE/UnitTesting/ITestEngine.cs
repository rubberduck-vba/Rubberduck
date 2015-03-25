using System;
using System.Collections.Generic;
using Rubberduck.Parsing;

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
        void Run(IEnumerable<TestMethod> tests);

        event EventHandler<TestCompleteEventArgs> TestComplete;
    }

    public class TestModuleEventArgs : EventArgs
    {
        private readonly string _projectName;
        private readonly string _moduleName;

        public TestModuleEventArgs(QualifiedModuleName qualifiedName)
            : this(qualifiedName.ProjectName, qualifiedName.ModuleName)
        {            
        }

        public TestModuleEventArgs(string projectName, string moduleName)
        {
            _projectName = projectName;
            _moduleName = moduleName;
        }

        public string ProjectName { get { return _projectName; } }
        public string ModuleName { get { return _moduleName; } }
    }

    public class TestMethodEventArgs : TestModuleEventArgs
    {
        private readonly string _memberName;

        public TestMethodEventArgs(QualifiedMemberName qualifiedName)
            : this(qualifiedName.QualifiedModuleName.ProjectName, qualifiedName.QualifiedModuleName.ModuleName, qualifiedName.Name)
        {
        }

        public TestMethodEventArgs(string projectName, string moduleName, string memberName)
            : base(projectName, moduleName)
        {
            _memberName = memberName;
        }

        public string MemberName { get { return _memberName; } 
        }
    }

    public class TestCompleteEventArgs : EventArgs
    {
        public TestResult Result { get; private set; }
        public TestMethod Test { get; private set; }

        public TestCompleteEventArgs(TestMethod test, TestResult result)
        {
            this.Test = test;
            this.Result = result;
        }
    }
}
