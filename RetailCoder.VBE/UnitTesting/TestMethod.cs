using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    public class TestMethod : IEquatable<TestMethod>
    {
        private readonly ICollection<TestResult> _assertResults = new List<TestResult>();
        private readonly IHostApplication _hostApp;

        public TestMethod(string projectName, string moduleName, string methodName, IHostApplication hostApp)
        {
            _projectName = projectName;
            _moduleName = moduleName;
            _methodName = methodName;
            _hostApp = hostApp;
        }

        private readonly string _projectName;
        public string ProjectName { get { return _projectName; } }

        private readonly string _moduleName;
        public string ModuleName { get { return _moduleName; } }

        private readonly string _methodName;
        public string MethodName { get { return _methodName; } }

        public string QualifiedName { get { return string.Concat(ProjectName, ".", ModuleName, ".", MethodName); } }

        public TestResult Run()
        {
            _assertResults.Clear(); //clear previous results to account for changes being made

            TestResult result;
            var duration = new TimeSpan();
            try
            {
                AssertHandler.OnAssertCompleted += HandleAssertCompleted;
                duration = _hostApp.TimedMethodCall(_projectName, _moduleName, _methodName);
                AssertHandler.OnAssertCompleted -= HandleAssertCompleted;
                
                result = EvaluateResults();
            }
            catch(Exception exception)
            {
                result = TestResult.Inconclusive("Test raised an error. " + exception.Message);
            }
            
            return new TestResult(result, duration.Milliseconds);
        }

        void HandleAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _assertResults.Add(e.Result);
        }

        private TestResult EvaluateResults()
        {
            var result = TestResult.Success();

            if (_assertResults.Any(assertion => assertion.Outcome == TestOutcome.Failed || assertion.Outcome == TestOutcome.Inconclusive))
            {
                result = _assertResults.First(assertion => assertion.Outcome == TestOutcome.Failed || assertion.Outcome == TestOutcome.Inconclusive);
            }

            return result;
        }

        public bool Equals(TestMethod other)
        {
            return QualifiedName == other.QualifiedName;
        }

        public override bool Equals(object obj)
        {
            return obj is TestMethod
                && ((TestMethod)obj).QualifiedName == QualifiedName;
        }

        public override int GetHashCode()
        {
            return QualifiedName.GetHashCode();
        }
    }
}
