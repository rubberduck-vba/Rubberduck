using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public class TestEngine : ITestEngine
    {
        public event EventHandler<TestCompletedEventArgs> TestComplete;

        public TestEngine()
        {
            AllTests = new Dictionary<TestMethod, TestResult>();
        }

        public IDictionary<TestMethod, TestResult> AllTests { get; set; }

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
        private void RunModuleInitialize(QualifiedModuleName qualifiedModuleName)
        {
            var handler = ModuleInitialize;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(qualifiedModuleName));
            }
        }

        public event EventHandler<TestModuleEventArgs> ModuleCleanup;
        private void RunModuleCleanup(QualifiedModuleName qualifiedModuleName)
        {
            var handler = ModuleCleanup;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(qualifiedModuleName));
            }
        }

        public event EventHandler<TestModuleEventArgs> MethodInitialize;
        private void RunMethodInitialize(QualifiedModuleName qualifiedModuleName)
        {
            var handler = MethodInitialize;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(qualifiedModuleName));
            }
        }

        public event EventHandler<TestModuleEventArgs> MethodCleanup;
        private void RunMethodCleanup(QualifiedModuleName qualifiedModuleName)
        {
            var handler = MethodCleanup;
            if (handler != null)
            {
                handler(this, new TestModuleEventArgs(qualifiedModuleName));
            }
        }

        public void Run(IEnumerable<TestMethod> tests, VBProject project)
        {
            //todo: move this to the "UI" layer. This code doesn't have to run for COM clients.
            //  COM clients will have to either already have a good reference, or be late bound.
            //  This is problematic for late bound code, because now we've *forced* them into early binding.
            project.EnsureReferenceToAddInLibrary();

            var testMethods = tests as IList<TestMethod> ?? tests.ToList();
            if (!testMethods.Any()) return;

            var methods = testMethods.ToDictionary(test => test, test => null as TestResult);
            AssignResults(methods.Keys);
        }

        private void AssignResults(IEnumerable<TestMethod> testMethods)
        {
            var tests = testMethods.ToList();

            var modules = tests.GroupBy(t => t.QualifiedMemberName.QualifiedModuleName);

            foreach (var module in modules)
            {
                RunModuleInitialize(module.Key);

                foreach (var test in module)
                {
                    if (tests.Contains(test))
                    {
                        RunMethodInitialize(test.QualifiedMemberName.QualifiedModuleName);
                    
                        var result = test.Run();
                        AllTests[test] = result;

                        RunMethodCleanup(test.QualifiedMemberName.QualifiedModuleName);


                        OnTestCompleted(new TestCompletedEventArgs(test, result));
                    }
                    else
                    {
                        AllTests[test] = null;
                    }
                }

                RunModuleCleanup(module.Key);
            }
        }

        protected virtual void OnTestCompleted(TestCompletedEventArgs arg)
        {
            if (TestComplete != null)
            {
                TestComplete(this, arg);
            }
        }
    }
}
