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
        /// <summary>   Runs VBA procedure specified by name. </summary>
        /// <param name="target"> Method Name of the Target VBA Procedure. </param>
        void Run(string target);

        /// <summary>   Timed call to Application.Run </summary>
        /// <param name="projectName">  Name of the project containing the method to be run. </param>
        /// <param name="moduleName">   Name of the module containing the method to be run. </param>
        /// <param name="methodName">   Name of the method run. </param>
        /// <returns>   Number of milliseconds it took to run the VBA procedure. </returns>
        long TimedMethodCall(string projectName, string moduleName, string methodName);
    }

    [ComVisible(false)]
    public abstract class HostApplicationBase<TApplication> : IHostApplication
    {
        protected readonly TApplication _application;
        protected HostApplicationBase(string applicationName)
        {
            _application = (TApplication)Marshal.GetActiveObject(applicationName + ".Application");
        }

        ~HostApplicationBase()
        {
            Marshal.ReleaseComObject(_application);
        }

        public abstract void Run(string target);
        protected abstract string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName);

        public long TimedMethodCall(string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(GenerateFullyQualifiedName(projectName, moduleName, methodName));

            stopwatch.Stop();
            return stopwatch.ElapsedMilliseconds;
        }
    }

    [ComVisible(false)]
    public class ExcelApp : HostApplicationBase<Excel.Application>
    {
        public ExcelApp() : base("Excel") { }

        public override void Run(string target)
        {
            base._application.Run(target);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(projectName, ".", moduleName, ".", methodName);
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
            //Access only supports Project.Procedure syntax. Error occurs if there are naming conflicts.
            // http://msdn.microsoft.com/en-us/library/office/ff193559(v=office.15).aspx
            // https://github.com/retailcoder/Rubberduck/issues/109
            _application.Run(string.Concat(projectName, ".", methodName));

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
