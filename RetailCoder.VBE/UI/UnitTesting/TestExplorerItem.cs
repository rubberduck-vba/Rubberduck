using System.Drawing;
using System.Runtime.InteropServices;
using Rubberduck.UnitTesting;

namespace Rubberduck.UI
{
    [ComVisible(false)]
    public class TestExplorerItem
    {
        public TestExplorerItem(TestMethod test, TestResult result)
        {
            _test = test;
            _result = result;
        }

        private readonly TestMethod _test;
        public TestMethod GetTestMethod()
        {
            return _test;
        }

        private TestResult _result;
        public void SetResult(TestResult result)
        {
            _result = result;
        }

        public Image Result { get { return _result.Icon(); } }
        public string ProjectName { get { return _test.ProjectName; } }
        public string ModuleName { get { return _test.ModuleName; } }
        public string MethodName { get { return _test.MethodName; } }
        public string Outcome { get { return _result == null ? string.Empty : _result.Outcome.ToString(); } }
        public string Message { get { return _result == null ? string.Empty : _result.Output; } }
        public string Duration { get { return _result == null ? string.Empty : _result.Duration.ToString() + " ms"; } }
    }
}