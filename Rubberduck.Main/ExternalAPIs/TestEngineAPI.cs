using Rubberduck.InternalApi.Common;
using Rubberduck.Resources.Registration;
using Rubberduck.UnitTesting;
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
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
        [Description("Runs all tests asynchronously and outputs the results to a file at the specified path")]
        [DispId(1)]
        void RunAllTestsAsync(string outputPath);
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
        /// Runs all unit tests writes the results to a file at the specified path.
        /// </summary>
        public void RunAllTestsAsync(string outputPath)
        {
            // Note that we can't use CanRun in case we are triggering the test via VBA since DesignMode is always set to false when you run a macro.
            // and CanRun interprets this as "not ready to run tests".

            Task.Run(() =>
            {
                var output = _testEngine.RunWithResults(_testEngine.Tests);
                if (!string.IsNullOrEmpty(outputPath))
                {
                    FileSystemProvider.FileSystem.File.WriteAllText(outputPath, output.ToString());
                }
            });
        }
    }

}
