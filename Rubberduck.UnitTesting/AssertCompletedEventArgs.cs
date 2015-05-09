﻿using System;

namespace Rubberduck.UnitTesting
{
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
