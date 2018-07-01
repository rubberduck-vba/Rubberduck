using System;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        //TestExplorerModel Model { get; }
        void Run();
        void Run(IEnumerable<TestMethod> tests);
        void Refresh();
        event EventHandler TestCompleted;
        ParserState[] AllowedRunStates { get; }
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
        // FIXME this needs to actually encapsulate the result as well
        public TestMethod Test { get; private set; }

        public TestCompletedEventArgs(TestMethod test)
        {
            Test = test;
        }
    }
}
