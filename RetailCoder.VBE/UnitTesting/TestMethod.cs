using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
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

        public string QualifiedName { get { return string.Concat(this.ProjectName, ".", this.ModuleName, ".", this.MethodName); } }

        public TestResult Run()
        {
            TestResult result;
            long duration = 0;
            try
            {
                AssertHandler.OnAssertCompleted += HandleAssertCompleted;
                duration =_hostApp.TimedMethodCall( _projectName, _moduleName, _methodName);
                AssertHandler.OnAssertCompleted -= HandleAssertCompleted;
                
                result = EvaluateResults();
            }
            catch(Exception exception)
            {
                result = TestResult.Inconclusive("Test raised an error. " + exception.Message);
            }
            
            return new TestResult(result, duration);
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
            return this.QualifiedName == other.QualifiedName;
        }

        public override bool Equals(object obj)
        {
            return obj is TestMethod
                && ((TestMethod)obj).QualifiedName == this.QualifiedName;
        }

        public override int GetHashCode()
        {
            return this.QualifiedName.GetHashCode();
        }
    }
}
