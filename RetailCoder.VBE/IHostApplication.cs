using System.Runtime.InteropServices;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
//using Access = Microsoft.Office.Interop.Access;
using Word = Microsoft.Office.Interop.Word;
using System;

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
        TimeSpan TimedMethodCall(string projectName, string moduleName, string methodName);
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

        public TimeSpan TimedMethodCall(string projectName, string moduleName, string methodName)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(GenerateFullyQualifiedName(projectName, moduleName, methodName));

            stopwatch.Stop();
            return stopwatch.Elapsed;
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

    //[ComVisible(false)]
    //public class AccessApp : HostApplicationBase<Access.Application>
    //{
    //    public AccessApp() : base("Access") { }

    //    public override void Run(string target)
    //    {
    //        base._application.Run(target);
    //    }

    //    protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
    //    {
    //        //Access only supports Project.Procedure syntax. Error occurs if there are naming conflicts.
    //        // http://msdn.microsoft.com/en-us/library/office/ff193559(v=office.15).aspx
    //        // https://github.com/retailcoder/Rubberduck/issues/109
            
    //        return string.Concat(projectName, ".", methodName);
    //    }
    //}

    [ComVisible(false)]
    public class WordApp : HostApplicationBase<Word.Application>
    {
        public WordApp() : base("Word") { }

        public override void Run(string target)
        {
            base._application.Run(target);
        }

        protected override string GenerateFullyQualifiedName(string projectName, string moduleName, string methodName)
        {
            return string.Concat(moduleName, ".", methodName);
        }
    }
}
