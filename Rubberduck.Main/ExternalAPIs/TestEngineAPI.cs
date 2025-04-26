using Rubberduck.InternalApi.Common;
using Rubberduck.Resources.Registration;
using Rubberduck.UnitTesting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.ExternalApi
{
    [
        ComVisible(true),
        Guid(RubberduckGuid.ITestEngineAPIInterfaceGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public interface ITestEngineAPI
    {
        [DispId(1)]
        string RunAllTestsAndGetResults(string filePath);
    }

    [
        ComVisible(true),
        Guid(RubberduckGuid.TestEngineAPIObjectGuid),
        ProgId(RubberduckProgId.TestEngineAPIProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(ITestEngineAPI)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public class TestEngineAPI : ITestEngineAPI
    {
        private readonly ITestEngine _testEngine;

        public TestEngineAPI(ITestEngine testEngine) 
        {
            _testEngine = testEngine;
        }

        /// <summary>
        /// Runs all unit tests and returns the results as a formatted string.
        /// </summary>
        /// <returns>A string containing the test results.</returns>
        public string RunAllTestsAndGetResults(string logPath)
        {

            // Note that we can't use CanRun in case we are triggering the test via VBA since DesignMode is always set to false when you run a macro.
            // and CanRun interprets this as "not ready to run tests".

            var task = Task.Run(() => {
                var output = _testEngine.RunWithResults(_testEngine.Tests);
                if (!string.IsNullOrEmpty(logPath))
                {
                    FileSystemProvider.FileSystem.File.WriteAllText(logPath, output.ToString());
                }
            });

            return "Task started to run tests asynchronously. Check the log file for results.";

        }
    }

}
