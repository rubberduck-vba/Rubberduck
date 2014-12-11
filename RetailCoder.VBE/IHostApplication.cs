using System.Runtime.InteropServices;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using Access = Microsoft.Office.Interop.Access;
using Word = Microsoft.Office.Interop.Word;

namespace Rubberduck
{
    [ComVisible(false)]
    public interface IHostApplication
    {
        void Run(string target);
        long TimedMethodCall(string projectName, string moduleName, string methodName);
    }

    [ComVisible(false)]
    public class ExcelApp : IHostApplication
    {
        Excel.Application _application;
        public ExcelApp()
        {
            _application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
        }

        ~ExcelApp()
        {
            Marshal.ReleaseComObject(_application);
        }

        public void Run(string target)
        {
            _application.Run(target);
        }

        /// <summary>   Timed call to Application.Run </summary>
        ///
        /// <param name="projectName">  Name of the project containing the method to be run. </param>
        /// <param name="moduleName">   Name of the module containing the method to be run. </param>
        /// <param name="methodName">   Name of the method run. </param>
        ///
        /// <returns>   Number of milliseconds it took to run the VBA procedure. </returns>
        public long TimedMethodCall(string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();
            //excel can have multiple projects, so it supports fully quantified names
            _application.Run(string.Concat(projectName, ".", moduleName, ".", methodName));

            stopwatch.Stop();

            return stopwatch.ElapsedMilliseconds;
        }
    }

    [ComVisible(false)]
    public class AccessApp : IHostApplication
    {
        Access.Application _application;
        public AccessApp()
        {
            _application = (Access.Application)Marshal.GetActiveObject("Access.Application");
        }

        ~AccessApp()
        {
            Marshal.ReleaseComObject(_application);
        }

        public void Run(string target)
        {
            _application.Run(target);
        }

        /// <summary>   Timed call to Application.Run </summary>
        ///
        /// <param name="projectName">  Name of the project containing the method to be run. </param>
        /// <param name="moduleName">   Name of the module containing the method to be run. </param>
        /// <param name="methodName">   Name of the method run. </param>
        ///
        /// <returns>   Number of milliseconds it took to run the VBA procedure. </returns>
        public long TimedMethodCall(string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();
            //access supports only single project; I would think it could handle Module.Procedure syntax, but it blows up with a com exception if I try.
            _application.Run(methodName);

            stopwatch.Stop();

            return stopwatch.ElapsedMilliseconds;
        }
    }

    [ComVisible(false)]
    public class WordApp : IHostApplication
    {
        Word.Application _application;
        public WordApp()
        {
            _application = (Word.Application)Marshal.GetActiveObject("Word.Application");
        }

        ~WordApp()
        {
            Marshal.ReleaseComObject(_application);
        }

        public void Run(string target)
        {
            _application.Run(target);
        }

        /// <summary>   Timed call to Application.Run </summary>
        ///
        /// <param name="projectName">  Name of the project containing the method to be run. </param>
        /// <param name="moduleName">   Name of the module containing the method to be run. </param>
        /// <param name="methodName">   Name of the method run. </param>
        ///
        /// <returns>   Number of milliseconds it took to run the VBA procedure. </returns>
        public long TimedMethodCall(string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();
            //Word supports single projects only
            _application.Run(string.Concat(moduleName, ".", methodName));

            stopwatch.Stop();

            return stopwatch.ElapsedMilliseconds;
        }
    }
}
