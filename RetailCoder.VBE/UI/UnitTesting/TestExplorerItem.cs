using System;
using System.Drawing;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.UnitTesting
{
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
        public string QualifiedMemberName { get { return _test.QualifiedMemberName.ToString(); } }
        public string ProjectName { get { return _test.QualifiedMemberName.QualifiedModuleName.Project.Name; } }
        public string ModuleName { get { return _test.QualifiedMemberName.QualifiedModuleName.Component.Name; } }
        public string MethodName { get { return _test.QualifiedMemberName.MemberName; } }
        public string Outcome { get { return _result == null ? string.Empty : _result.Outcome.ToString(); } }
        public string Message { get { return _result == null ? string.Empty : _result.Output; } }
        public string Duration { get { return _result == null ? string.Empty : _result.Duration + " ms"; } }

        public TimeSpan GetDuration()
        {
            return _result == null ? new TimeSpan() : TimeSpan.FromMilliseconds(_result.Duration);
        }
    }
}