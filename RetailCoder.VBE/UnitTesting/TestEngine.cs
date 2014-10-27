using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using RetailCoderVBE.Reflection;
using System;

namespace RetailCoderVBE.UnitTesting
{
    [ComVisible(true)]
    internal interface ITestOutput
    {
        void WriteResult(TestMethod method, TestResult result);
    }

    [ComVisible(false)]
    internal class TestEngine
    {
        public TestEngine(ITestOutput output)
        {
            _output = output;
        }

        private readonly ITestOutput _output;

        /// <summary>
        /// Runs specified test.
        /// </summary>
        /// <param name="test"></param>
        public TestResult Run(TestMethod test)
        {
            var result = test.Run();
            _output.WriteResult(test, result);

            return result;
        }
    }
}
