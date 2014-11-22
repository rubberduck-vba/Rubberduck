using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Rubberduck.UnitTesting
{
    [ComVisible(false)]
    public static class ExcelApplicationExtensions
    {
        public static long TimedMethodCall(this Microsoft.Office.Interop.Excel.Application application, string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();

            application.Run(string.Concat(projectName, ".", moduleName, ".", methodName));
            stopwatch.Stop();
            
            return stopwatch.ElapsedMilliseconds;
        }
    }
}