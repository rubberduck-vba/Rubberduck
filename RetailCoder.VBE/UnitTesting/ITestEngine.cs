using System;
using System.Collections.Generic;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        TestExplorerModel Model { get; }
        void Run();
        void Run(IEnumerable<TestMethod> tests);
        void Refresh();
        event EventHandler TestCompleted;
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

    public class TestCompletedEventArgs : EventArgs
    {
        public TestMethod Test { get; private set; }

        public TestCompletedEventArgs(TestMethod test)
        {
            Test = test;
        }
    }
}
