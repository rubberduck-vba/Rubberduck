using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UnitTesting
{
    [ComVisible(true)]
    public interface ITestRunner
    {
        /// <summary>
        /// Runs all Rubberduck unit tests in the IDE, optionally outputting results to specified text file.
        /// </summary>
        /// <param name="vbe"></param>
        /// <param name="outputFilePath"></param>
        /// <returns>Returns a string containing the test results.</returns>
        string RunAllTests(VBE vbe, string outputFilePath = null);
    }
}