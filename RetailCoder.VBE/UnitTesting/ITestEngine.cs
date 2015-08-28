using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public interface ITestEngine
    {
        event EventHandler<TestModuleEventArgs> ModuleInitialize;
        event EventHandler<TestModuleEventArgs> ModuleCleanup;
        event EventHandler<TestModuleEventArgs> MethodInitialize;
        event EventHandler<TestModuleEventArgs> MethodCleanup;
        TestExplorerModelBase Model { get; }
        void Run();
        void Run(IEnumerable<TestMethod> tests);

        event EventHandler<TestCompletedEventArgs> TestComplete;
    }

    public class TestEngine : ITestEngine
    {
        private readonly TestExplorerModelBase _model;

        public TestEngine(TestExplorerModelBase model)
        {
            _model = model;
        }

        // todo: remove these events after confirming that they're not needed
        public event EventHandler<TestModuleEventArgs> ModuleInitialize;
        public event EventHandler<TestModuleEventArgs> ModuleCleanup;
        public event EventHandler<TestModuleEventArgs> MethodInitialize;
        public event EventHandler<TestModuleEventArgs> MethodCleanup;
        public event EventHandler<TestCompletedEventArgs> TestComplete;

        private void RaiseEvent<T>(EventHandler<T> method, T args)
        {
            var handler = method;
            if (handler != null)
            {
                handler.Invoke(this, args);
            }
        }

        public TestExplorerModelBase Model { get { return _model; } }

        public void Run()
        {
            _model.Refresh();
            Run(_model.Tests);
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            var testMethods = tests as IList<TestMethod> ?? tests.ToList();
            if (!testMethods.Any())
            {
                return;
            }

            var modules = testMethods.GroupBy(test => test.QualifiedMemberName.QualifiedModuleName);
            foreach (var module in modules)
            {
                var moduleEventArgs = new TestModuleEventArgs(module.Key);
                RaiseEvent(ModuleInitialize, moduleEventArgs);

                foreach (var test in module)
                {
                    RaiseEvent(MethodInitialize, moduleEventArgs);

                    test.Run();

                    RaiseEvent(MethodCleanup, moduleEventArgs);

                    var completedEventArgs = new TestCompletedEventArgs(test);
                    RaiseEvent(TestComplete, completedEventArgs);
                }

                RaiseEvent(ModuleCleanup, moduleEventArgs);
            }
        }
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
