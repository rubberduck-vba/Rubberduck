using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.VBEHost;

namespace Rubberduck.UnitTesting
{
    public class TestMethod : IEquatable<TestMethod>
    {
        private readonly ICollection<TestResult> _assertResults = new List<TestResult>();
        private readonly IHostApplication _hostApp;

        public TestMethod(QualifiedMemberName qualifiedMemberName, IHostApplication hostApp)
        {
            _qualifiedMemberName = qualifiedMemberName;
            _hostApp = hostApp;
        }

        private readonly QualifiedMemberName _qualifiedMemberName;
        public QualifiedMemberName QualifiedMemberName { get { return _qualifiedMemberName; } }

        public TestResult Run()
        {
            _assertResults.Clear(); //clear previous results to account for changes being made

            TestResult result;
            var duration = new TimeSpan();
            try
            {
                AssertHandler.OnAssertCompleted += HandleAssertCompleted;
                duration = _hostApp.TimedMethodCall(_qualifiedMemberName);
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
            return QualifiedMemberName.Equals(other.QualifiedMemberName);
        }

        public override bool Equals(object obj)
        {
            return obj is TestMethod
                && ((TestMethod)obj).QualifiedMemberName.Equals(QualifiedMemberName);
        }

        public override int GetHashCode()
        {
            return QualifiedMemberName.GetHashCode();
        }
    }
}
