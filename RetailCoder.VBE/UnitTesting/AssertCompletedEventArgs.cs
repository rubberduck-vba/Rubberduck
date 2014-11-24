using System;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public class AssertCompletedEventArgs : EventArgs
    {
        public AssertCompletedEventArgs(TestResult result)
        {
            _result = result;
        }

        private readonly TestResult _result;
        public TestResult Result { get { return _result; } }
    }
}
