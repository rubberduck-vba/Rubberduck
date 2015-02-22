using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UnitTesting
{
    public class TestEngine2 : ITestEngine
    {
        private IEnumerable<TestMethod> _lastRun;

        public event EventHandler<TestCompleteEventArg> TestComplete;
        public event EventHandler<EventArgs> AllTestsComplete;

        public TestEngine2()
        {
            this.AllTests = new Dictionary<TestMethod, TestResult>();
        }

        void TestEngine2_AllTestsComplete(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        public IDictionary<TestMethod, TestResult> AllTests
        {
            get; set;
        }

        public IEnumerable<TestMethod> FailedTests()
        {
                return this.AllTests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Failed)
                                    .Select(test => test.Key);
        }

        public IEnumerable<TestMethod> LastRunTests(TestOutcome? outcome = null)
        {
            return this.AllTests.Where(test => test.Value != null
                 && test.Value.Outcome == (outcome ?? test.Value.Outcome))
                 .Select(test => test.Key);
        }

        public IEnumerable<TestMethod> NotRunTests()
        {
                return this.AllTests.Where(test => test.Value == null)
                                    .Select(test => test.Key);
        }

        public IEnumerable<TestMethod> PassedTests()
        {
                return this.AllTests.Where(test => test.Value != null && test.Value.Outcome == TestOutcome.Succeeded)
                                    .Select(test => test.Key);
        }

        public void ReRun()
        {
            //todo: implement or remove
            throw new NotImplementedException();
        }

        public void Run()
        {
            if (!this.AllTests.Any())
            {
                _lastRun = null;
                return;
            }

            Run(this.AllTests.Keys);
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            if (tests.Any())
            {
                var methods = tests.ToDictionary(test => test, test => null as TestResult);
                AssignResults(methods.Keys);
                _lastRun = methods.Keys;
            }
            else
            {
                _lastRun = null;
            }
        }

        private void AssignResults(IEnumerable<TestMethod> testMethods)
        {
            var tests = testMethods.ToList();
            var keys = this.AllTests.Keys.ToList();
            foreach (var test in keys)
            {
                if (tests.Contains(test))
                {
                    var result = test.Run();
                    this.AllTests[test] = result;
                    OnTestComplete(new TestCompleteEventArg(test, result));
                }
                else
                {
                    this.AllTests[test] = null;
                }
            }
        }

        protected virtual void OnTestComplete(TestCompleteEventArg arg)
        {
            TestComplete(this, arg);
        }
    }
}
