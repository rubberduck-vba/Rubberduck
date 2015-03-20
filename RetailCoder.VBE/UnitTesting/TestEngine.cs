using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UnitTesting
{
    public class TestEngine : ITestEngine
    {
        public event EventHandler<TestCompleteEventArgs> TestComplete;

        public TestEngine()
        {
            AllTests = new Dictionary<TestMethod, TestResult>();
        }

        public IDictionary<TestMethod, TestResult> AllTests
        {
            get;
            set;
        }

        public IEnumerable<TestMethod> FailedTests()
        {
            return AllTests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Failed)
                                .Select(test => test.Key);
        }

        public IEnumerable<TestMethod> LastRunTests(TestOutcome? outcome = null)
        {
            return AllTests.Where(test => test.Value != null
                 && test.Value.Outcome == (outcome ?? test.Value.Outcome))
                 .Select(test => test.Key);
        }

        public IEnumerable<TestMethod> NotRunTests()
        {
            return AllTests.Where(test => test.Value == null)
                                .Select(test => test.Key);
        }

        public IEnumerable<TestMethod> PassedTests()
        {
            return AllTests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Succeeded)
                                .Select(test => test.Key);
        }

        public event EventHandler<TestModuleEventArgs> ModuleInitialize;
        private void RunModuleInitialize(string projectName, string moduleName)
        {
            var handler = ModuleInitialize;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(projectName, moduleName));
            }
        }

        public event EventHandler<TestModuleEventArgs> ModuleCleanup;
        private void RunModuleCleanup(string projectName, string moduleName)
        {
            var handler = ModuleCleanup;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(projectName, moduleName));
            }
        }

        public event EventHandler<TestModuleEventArgs> MethodInitialize;
        private void RunMethodInitialize(string projectName, string moduleName)
        {
            var handler = MethodInitialize;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(projectName, moduleName));
            }
        }

        public event EventHandler<TestModuleEventArgs> MethodCleanup;
        private void RunMethodCleanup(string projectName, string moduleName)
        {
            var handler = MethodCleanup;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(projectName, moduleName));
            }
        }

        public void Run()
        {
            Run(AllTests.Keys);
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            if (tests.Any())
            {
                var methods = tests.ToDictionary(test => test, test => null as TestResult);
                AssignResults(methods.Keys);
            }
        }

        private void AssignResults(IEnumerable<TestMethod> testMethods)
        {
            var tests = testMethods.ToList();

            var modules = tests.GroupBy(t => new {Project = t.ProjectName, Module = t.ModuleName});

            foreach (var module in modules)
            {
                RunModuleInitialize(module.Key.Project, module.Key.Module);

                foreach (var test in module)
                {
                    if (tests.Contains(test))
                    {
                        RunMethodInitialize(test.ProjectName, test.ModuleName);
                    
                        var result = test.Run();
                        AllTests[test] = result;

                        RunMethodCleanup(test.ProjectName, test.ModuleName);


                        OnTestComplete(new TestCompleteEventArgs(test, result));
                    }
                    else
                    {
                        AllTests[test] = null;
                    }
                }

                RunModuleCleanup(module.Key.Project, module.Key.Module);
            }
        }

        protected virtual void OnTestComplete(TestCompleteEventArgs arg)
        {
            if (TestComplete != null)
            {
                TestComplete(this, arg);
            }
        }
    }
}
